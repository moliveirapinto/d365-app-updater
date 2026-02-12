// Global variables
let msalInstance = null;
let ppToken = null; // Power Platform API token
let environmentId = null;
let currentOrgUrl = null;
let apps = [];
let allEnvironments = []; // Cached list of all environments
let selectedApps = new Set(); // Multi-select tracking

// ═══════════════════════════════════════════════════════════════════════════
// LOGGING SYSTEM - Persists across redirects for debugging auth flows
// ═══════════════════════════════════════════════════════════════════════════
const LOG_STORAGE_KEY = 'd365_app_logs';
const MAX_LOGS = 200;

function appLog(level, message, data = null) {
    const timestamp = new Date().toISOString();
    const logEntry = { timestamp, level, message, data: data ? JSON.stringify(data) : null };
    
    // Console output
    const consoleMsg = `[${timestamp}] [${level}] ${message}`;
    if (level === 'ERROR') {
        console.error(consoleMsg, data || '');
    } else if (level === 'WARN') {
        console.warn(consoleMsg, data || '');
    } else {
        console.log(consoleMsg, data || '');
    }
    
    // Persist to sessionStorage (survives redirects within same session)
    try {
        const logs = JSON.parse(sessionStorage.getItem(LOG_STORAGE_KEY) || '[]');
        logs.push(logEntry);
        // Keep only last MAX_LOGS entries
        while (logs.length > MAX_LOGS) logs.shift();
        sessionStorage.setItem(LOG_STORAGE_KEY, JSON.stringify(logs));
    } catch (e) {
        console.error('Failed to persist log:', e);
    }
    
    // Update UI log panel if visible
    updateLogPanel();
}

function updateLogPanel() {
    const panel = document.getElementById('logPanel');
    const content = document.getElementById('logContent');
    if (!panel || !content) return;
    
    try {
        const logs = JSON.parse(sessionStorage.getItem(LOG_STORAGE_KEY) || '[]');
        content.innerHTML = logs.map(log => {
            const color = log.level === 'ERROR' ? '#ff4444' : log.level === 'WARN' ? '#ffaa00' : '#00cc00';
            const time = log.timestamp.split('T')[1].split('.')[0];
            const dataStr = log.data ? `\n    └─ ${log.data}` : '';
            return `<div style="color:${color};margin:2px 0;"><span style="color:#888">[${time}]</span> <b>[${log.level}]</b> ${log.message}${dataStr}</div>`;
        }).join('');
        content.scrollTop = content.scrollHeight;
    } catch (e) {}
}

function clearLogs() {
    sessionStorage.removeItem(LOG_STORAGE_KEY);
    updateLogPanel();
    appLog('INFO', 'Logs cleared');
}

function exportLogs() {
    try {
        const logs = JSON.parse(sessionStorage.getItem(LOG_STORAGE_KEY) || '[]');
        const text = logs.map(l => `[${l.timestamp}] [${l.level}] ${l.message}${l.data ? ' | ' + l.data : ''}`).join('\n');
        const blob = new Blob([text], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `d365-app-logs-${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
        a.click();
        URL.revokeObjectURL(url);
    } catch (e) {
        alert('Failed to export logs: ' + e.message);
    }
}

function toggleLogPanel() {
    const panel = document.getElementById('logPanel');
    if (panel) {
        const isHidden = panel.style.display === 'none';
        panel.style.display = isHidden ? 'block' : 'none';
        if (isHidden) updateLogPanel();
    }
}

// Shorthand logging functions
const logInfo = (msg, data) => appLog('INFO', msg, data);
const logWarn = (msg, data) => appLog('WARN', msg, data);
const logError = (msg, data) => appLog('ERROR', msg, data);
const logDebug = (msg, data) => appLog('DEBUG', msg, data);

// ═══════════════════════════════════════════════════════════════════════════

// Supabase config for usage tracking
const SUPABASE_URL = 'https://fpekzltxukikaixebeeu.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZwZWt6bHR4dWtpa2FpeGViZWV1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzA0MDU0ODEsImV4cCI6MjA4NTk4MTQ4MX0.uH4JgKbf_-Al_iArzEy6UZ3edJNzFSCBVlMNI04li0Y';

// MSAL Configuration
function createMsalConfig(tenantId, clientId) {
    // Compute redirect URI from current path (handles GitHub Pages subpaths like /d365-app-updater/)
    const pathDir = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
    // Use 'organizations' endpoint if no tenant ID provided (allows any Azure AD account)
    const authority = tenantId 
        ? `https://login.microsoftonline.com/${tenantId}`
        : 'https://login.microsoftonline.com/organizations';
    return {
        auth: {
            clientId: clientId,
            authority: authority,
            redirectUri: window.location.origin + pathDir,
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false,
        },
    };
}

// Flag to track if we're resuming from a redirect
let _pendingRedirectAuth = false;

// Initialize on page load
document.addEventListener('DOMContentLoaded', async function() {
    // ═══ EMERGENCY RESET: Add ?reset to URL to clear all auth state ═══
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.has('reset')) {
        console.log('🔄 RESET MODE: Clearing all authentication state...');
        localStorage.removeItem('d365_app_updater_creds');
        sessionStorage.removeItem('d365_auth_step');
        sessionStorage.removeItem('d365_redirect_count');
        sessionStorage.removeItem(LOG_STORAGE_KEY);
        // Clear MSAL cache
        localStorage.removeItem('msal.token.keys.' + urlParams.get('reset'));
        // Clear all MSAL-related items
        Object.keys(localStorage).forEach(key => {
            if (key.startsWith('msal.')) localStorage.removeItem(key);
        });
        Object.keys(sessionStorage).forEach(key => {
            if (key.startsWith('msal.') || key.startsWith('d365_') || key.startsWith('wizard_')) {
                sessionStorage.removeItem(key);
            }
        });
        alert('✅ All authentication data cleared! You can now start fresh.');
        // Redirect to clean URL
        window.location.href = window.location.origin + window.location.pathname;
        return;
    }

    logInfo('=== APP INITIALIZATION ===');
    logInfo('URL', { href: window.location.href, hash: window.location.hash ? 'present' : 'none', pathname: window.location.pathname });
    logInfo('Auth step in storage', sessionStorage.getItem('d365_auth_step'));
    logInfo('Saved creds', { 
        sessionStorage: !!sessionStorage.getItem('d365_app_updater_creds_temp'), 
        localStorage: !!localStorage.getItem('d365_app_updater_creds') 
    });

    // If returning from wizard auth redirect, handle MSAL here on the root page
    // (MSAL requires redirect response to be processed on the same URL that was used as redirectUri)
    const wizardClientId = sessionStorage.getItem('wizard_clientId');
    if (wizardClientId) {
        logInfo('Wizard flow detected', { wizardClientId: wizardClientId.substring(0, 8) + '...', hasHash: !!window.location.hash });
        
        // First check if the hash contains an error
        if (window.location.hash && window.location.hash.includes('error')) {
            const hashParams = new URLSearchParams(window.location.hash.substring(1));
            const error = hashParams.get('error');
            const errorDesc = decodeURIComponent(hashParams.get('error_description') || 'Unknown error');
            logError('Wizard auth error in URL', { error, errorDesc });
            sessionStorage.setItem('wizard_error', errorDesc);
            sessionStorage.removeItem('wizard_clientId');
            window.location.replace('setup-wizard.html');
            return;
        }
        
        // Only process if we have a hash (returning from redirect)
        if (window.location.hash) {
            logInfo('Processing wizard redirect...');
            try {
                const pathDir = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
                const wizardMsalConfig = {
                    auth: {
                        clientId: wizardClientId,
                        authority: 'https://login.microsoftonline.com/organizations',
                        redirectUri: window.location.origin + pathDir
                    },
                    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
                };
                const wizardMsal = new msal.PublicClientApplication(wizardMsalConfig);
                await wizardMsal.initialize();
                logInfo('Wizard MSAL initialized, calling handleRedirectPromise...');
                
                const response = await wizardMsal.handleRedirectPromise();
                logInfo('Wizard handleRedirectPromise result', response ? { hasToken: !!response.accessToken, scopes: response.scopes } : 'null');
                
                if (response && response.accessToken) {
                    sessionStorage.setItem('wizard_accessToken', response.accessToken);
                    logInfo('Wizard token saved to sessionStorage');
                } else {
                    // No response from redirect, try to get token silently if there's an account
                    const accounts = wizardMsal.getAllAccounts();
                    logInfo('Wizard accounts', accounts.length);
                    if (accounts.length > 0) {
                        try {
                            const silentResult = await wizardMsal.acquireTokenSilent({
                                scopes: ['https://graph.microsoft.com/Application.ReadWrite.All', 'https://graph.microsoft.com/DelegatedPermissionGrant.ReadWrite.All'],
                                account: accounts[0]
                            });
                            if (silentResult && silentResult.accessToken) {
                                sessionStorage.setItem('wizard_accessToken', silentResult.accessToken);
                                logInfo('Wizard token acquired silently');
                            }
                        } catch (silentErr) {
                            logWarn('Wizard silent token acquisition failed', silentErr.message);
                        }
                    }
                }
                
                // Check if we got a token
                if (!sessionStorage.getItem('wizard_accessToken')) {
                    logError('No wizard token acquired after redirect processing');
                    sessionStorage.setItem('wizard_error', 'Failed to acquire access token. Please ensure you have admin permissions and try again.');
                }
            } catch (err) {
                logError('Wizard redirect handling error', err.message);
                sessionStorage.setItem('wizard_error', err.message);
            }
            // Forward to setup-wizard.html (without the hash)
            window.location.replace('setup-wizard.html');
            return;
        } else {
            // wizardClientId is set but no hash - user might be stuck, clear and let them restart
            logWarn('Wizard clientId set but no hash, clearing wizard state');
            sessionStorage.removeItem('wizard_clientId');
        }
    }

    hideLoading();
    
    if (typeof msal === 'undefined') {
        logError('MSAL library failed to load');
        alert('Error: MSAL library failed to load.');
        return;
    }
    logInfo('MSAL library loaded successfully');
    
    const redirectUriElement = document.getElementById('redirectUri');
    if (redirectUriElement) {
        redirectUriElement.textContent = window.location.origin;
    }
    
    loadSavedCredentials();
    
    document.getElementById('authForm').addEventListener('submit', handleAuthentication);
    document.getElementById('logoutBtn').addEventListener('click', handleLogout);
    document.getElementById('refreshAppsBtn').addEventListener('click', loadApplications);
    document.getElementById('updateAllBtn').addEventListener('click', updateAllApps);
    document.getElementById('reinstallAllBtn').addEventListener('click', reinstallAllApps);
    document.getElementById('updateSelectedBtn').addEventListener('click', updateSelectedApps);
    
    // Close environment dropdown when clicking outside
    document.addEventListener('click', function(e) {
        const switcher = document.querySelector('.env-switcher');
        if (switcher && !switcher.contains(e.target)) {
            closeEnvDropdown();
        }
    });
    
    // Try to handle redirect response first, then fall back to auto-login
    handleRedirectResponse().then(() => {
        console.log('App initialized');
    });
});

// Load saved credentials
function loadSavedCredentials() {
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            const orgUrlEl = document.getElementById('orgUrl');
            if (orgUrlEl) orgUrlEl.value = creds.orgUrl || creds.organizationId || creds.environmentId || '';
            document.getElementById('clientId').value = creds.clientId || '';
            document.getElementById('rememberMe').checked = true;
        } catch (e) {}
    }
}

// Handle redirect response when returning from Microsoft login redirect
async function handleRedirectResponse() {
    logInfo('=== handleRedirectResponse START ===');
    
    // Check if URL hash contains an error response from Azure AD
    const hash = window.location.hash;
    if (hash && (hash.includes('error=') || hash.includes('error_description='))) {
        logError('Azure AD returned an error in the URL');
        
        // Try to decode the error
        const hashParams = new URLSearchParams(hash.substring(1));
        const error = hashParams.get('error');
        const errorDesc = decodeURIComponent(hashParams.get('error_description') || 'Unknown error');
        
        logError('Azure AD Error', { error, errorDesc });
        
        // Clear all auth state
        localStorage.removeItem('d365_app_updater_creds');
        sessionStorage.removeItem('d365_auth_step');
        sessionStorage.removeItem('d365_redirect_count');
        Object.keys(localStorage).forEach(key => {
            if (key.startsWith('msal.')) localStorage.removeItem(key);
        });
        Object.keys(sessionStorage).forEach(key => {
            if (key.startsWith('msal.')) sessionStorage.removeItem(key);
        });
        
        // Clean the URL
        history.replaceState(null, '', window.location.pathname);
        
        // Show error
        const resetButton = `<br><br><button onclick="window.location.reload()" style="background:#0078d4;color:white;border:none;padding:10px 20px;border-radius:6px;cursor:pointer;font-weight:600;">🔄 Start Fresh</button>`;
        
        let friendlyMessage = errorDesc;
        if (errorDesc.includes('AADSTS650057') || errorDesc.includes('not listed in the requested permissions')) {
            friendlyMessage = `<strong>Missing API Permissions</strong><br><br>
Your Azure AD app registration is missing required permissions.<br><br>
<strong>To fix this:</strong><br>
1. Go to <a href="https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">Azure Portal → App Registrations</a><br>
2. Find and click on your app<br>
3. Go to <strong>API permissions</strong> → <strong>Add a permission</strong><br>
4. Add: <strong>Power Platform API</strong> → Delegated → <code>user_impersonation</code><br>
5. Add: <strong>Dynamics CRM</strong> → Delegated → <code>user_impersonation</code><br>
6. Click <strong>Grant admin consent</strong>${resetButton}`;
        } else {
            friendlyMessage = `<strong>Authentication Error</strong><br><br>${errorDesc}${resetButton}`;
        }
        
        showError(friendlyMessage);
        return;
    }
    
    // Try to get credentials from sessionStorage first (survives redirect, works with tracking prevention)
    // then fall back to localStorage (for "remember me" functionality)
    let savedCreds = sessionStorage.getItem('d365_app_updater_creds_temp');
    let credsSource = 'sessionStorage';
    
    if (!savedCreds) {
        savedCreds = localStorage.getItem('d365_app_updater_creds');
        credsSource = 'localStorage';
    }
    
    if (!savedCreds) {
        logInfo('No saved credentials in sessionStorage or localStorage, user must log in manually');
        return;
    }
    
    logInfo('Found credentials in ' + credsSource);

    let creds;
    try {
        creds = JSON.parse(savedCreds);
        logDebug('Parsed saved credentials', { orgUrl: creds.orgUrl, tenantId: creds.tenantId?.substring(0,8) + '...', hasClientId: !!creds.clientId });
    } catch (e) {
        logError('Failed to parse saved credentials', e.message);
        return;
    }

    const orgUrlValue = creds.orgUrl || creds.organizationId || creds.environmentId || '';
    const tenantId = creds.tenantId || '';
    const clientId = creds.clientId || '';
    
    // Tenant ID is optional - only Client ID and Org URL are required
    if (!clientId || !orgUrlValue) {
        logWarn('Missing required credentials', { hasOrgUrl: !!orgUrlValue, hasTenantId: !!tenantId, hasClientId: !!clientId });
        return;
    }

    try {
        logInfo('Creating MSAL instance...');
        const msalConfig = createMsalConfig(tenantId, clientId);
        logDebug('MSAL config', { redirectUri: msalConfig.auth.redirectUri, authority: msalConfig.auth.authority });
        
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        logInfo('MSAL initialized');

        // Check if we're returning from a redirect login
        logInfo('Calling handleRedirectPromise...');
        const redirectResult = await msalInstance.handleRedirectPromise();
        logInfo('handleRedirectPromise result', redirectResult ? { 
            hasToken: !!redirectResult.accessToken, 
            scopes: redirectResult.scopes,
            account: redirectResult.account?.username 
        } : 'null');

        const accounts = msalInstance.getAllAccounts();
        logInfo('MSAL accounts found', { count: accounts.length, accounts: accounts.map(a => a.username) });
        
        if (accounts.length === 0) {
            logInfo('No accounts in cache, user must log in manually');
            msalInstance = null;
            sessionStorage.removeItem('d365_auth_step');
            return;
        }

        const account = accounts[0];
        logInfo('Using account', account.username);

        // Track which step of auth we're on to avoid redirect loops
        const authStep = sessionStorage.getItem('d365_auth_step') || 'none';
        logInfo('Current auth step', authStep);
        
        // Check redirect counter to prevent infinite loops
        let redirectCount = parseInt(sessionStorage.getItem('d365_redirect_count') || '0', 10);
        logInfo('Redirect count', redirectCount);
        
        if (redirectCount > 5) {
            logError('Too many redirects, breaking loop');
            sessionStorage.removeItem('d365_auth_step');
            sessionStorage.removeItem('d365_redirect_count');
            throw new Error('Authentication failed after multiple attempts. Please clear your browser cache and try again.');
        }

        // If returning from a runtime BAP token redirect, clear step and continue
        if (authStep === 'acquiring_bap_runtime' && redirectResult) {
            logInfo('Returned from runtime BAP token redirect, clearing step');
            sessionStorage.removeItem('d365_auth_step');
        }

        // If we just came back from login (initial or login_redirect step), the redirectResult is for the login
        // We need to proceed to acquire PP and BAP tokens
        if ((authStep === 'login_redirect' || authStep === 'initial') && redirectResult) {
            logInfo('Returned from initial login redirect, proceeding to acquire API tokens');
            sessionStorage.removeItem('d365_auth_step'); // Clear so we can proceed
        }

        showLoading('Authenticating...', 'Restoring your session');

        // ─── Acquire Power Platform token ───────────────────────────────────────
        logInfo('Acquiring Power Platform token...');
        const ppRequest = { scopes: ['https://api.powerplatform.com/.default'], account };
        let ppResult;
        
        // Check if redirect result contains PP scope
        const isPPRedirectResult = redirectResult && redirectResult.scopes && 
            redirectResult.scopes.some(s => s.includes('api.powerplatform.com'));
        
        if (authStep === 'acquiring_pp' && isPPRedirectResult && redirectResult.accessToken) {
            ppResult = redirectResult;
            logInfo('Using PP token from redirect result', { scopes: redirectResult.scopes });
            sessionStorage.removeItem('d365_auth_step');
            sessionStorage.removeItem('d365_redirect_count');
        } else {
            try {
                logDebug('Trying acquireTokenSilent for PP...');
                ppResult = await msalInstance.acquireTokenSilent(ppRequest);
                logInfo('PP token acquired silently');
            } catch (e) {
                logWarn('acquireTokenSilent for PP failed', { error: e.message, errorCode: e.errorCode });
                // Only redirect if we haven't tried too many times
                if (authStep !== 'acquiring_pp') {
                    logInfo('Redirecting for PP token consent...');
                    sessionStorage.setItem('d365_auth_step', 'acquiring_pp');
                    sessionStorage.setItem('d365_redirect_count', String(redirectCount + 1));
                    await msalInstance.acquireTokenRedirect(ppRequest);
                    return;
                } else {
                    logError('Already tried PP redirect, still failing', e.message);
                    sessionStorage.removeItem('d365_auth_step');
                    sessionStorage.removeItem('d365_redirect_count');
                    throw new Error('Failed to acquire Power Platform token. Error: ' + e.message + '. Please check your app registration has the correct API permissions.');
                }
            }
        }
        ppToken = ppResult.accessToken;
        logInfo('PP token acquired successfully');

        // ─── Acquire BAP token ──────────────────────────────────────────────────
        logInfo('Acquiring BAP token...');
        const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account };
        
        // Check if redirect result contains BAP scope
        const isBAPRedirectResult = redirectResult && redirectResult.scopes && 
            redirectResult.scopes.some(s => s.includes('api.bap.microsoft.com'));
        
        if (authStep === 'acquiring_bap' && isBAPRedirectResult && redirectResult.accessToken) {
            logInfo('Using BAP token from redirect result', { scopes: redirectResult.scopes });
            sessionStorage.removeItem('d365_auth_step');
            sessionStorage.removeItem('d365_redirect_count');
        } else {
            try {
                logDebug('Trying acquireTokenSilent for BAP...');
                await msalInstance.acquireTokenSilent(bapRequest);
                logInfo('BAP token acquired silently');
            } catch (e) {
                logWarn('acquireTokenSilent for BAP failed', { error: e.message, errorCode: e.errorCode });
                if (authStep !== 'acquiring_bap') {
                    logInfo('Redirecting for BAP token consent...');
                    sessionStorage.setItem('d365_auth_step', 'acquiring_bap');
                    sessionStorage.setItem('d365_redirect_count', String(redirectCount + 1));
                    await msalInstance.acquireTokenRedirect(bapRequest);
                    return;
                } else {
                    logError('Already tried BAP redirect, still failing', e.message);
                    sessionStorage.removeItem('d365_auth_step');
                    sessionStorage.removeItem('d365_redirect_count');
                    throw new Error('Failed to acquire BAP token. Error: ' + e.message + '. Please check your app registration has the correct API permissions.');
                }
            }
        }

        // Clear auth step - we're done with redirects
        sessionStorage.removeItem('d365_auth_step');
        logInfo('Auth step cleared, proceeding to resolve environment');

        showLoading('Authenticating...', 'Resolving environment');

        // Normalize org URL
        let normalizedOrgUrl = orgUrlValue;
        if (!normalizedOrgUrl.startsWith('https://')) {
            normalizedOrgUrl = 'https://' + normalizedOrgUrl;
        }
        normalizedOrgUrl = normalizedOrgUrl.replace(/\/+$/, '');
        logInfo('Resolving environment for URL', normalizedOrgUrl);

        environmentId = await resolveOrgUrlToEnvironmentId(normalizedOrgUrl);
        if (!environmentId) {
            throw new Error('Could not resolve environment. Please verify the Organization URL and your permissions.');
        }
        logInfo('Environment resolved', environmentId);

        currentOrgUrl = normalizedOrgUrl;

        showLoading('Authenticating...', 'Loading environment details');
        await getEnvironmentName();

        hideLoading();

        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appsSection').classList.remove('hidden');

        // Load schedule settings
        loadSchedule();

        // Clean up temp credentials from sessionStorage
        sessionStorage.removeItem('d365_app_updater_creds_temp');
        
        // If user wanted to remember, try to save to localStorage (may fail with tracking prevention)
        if (creds.rememberMe) {
            try {
                localStorage.setItem('d365_app_updater_creds', JSON.stringify({ orgUrl: creds.orgUrl, tenantId: creds.tenantId, clientId: creds.clientId }));
            } catch (e) {
                logWarn('Could not save to localStorage for remember me', e.message);
            }
        }

        logInfo('=== AUTH SUCCESS ===', account.username);
        await loadApplications();

    } catch (e) {
        sessionStorage.removeItem('d365_auth_step');
        sessionStorage.removeItem('d365_redirect_count');
        sessionStorage.removeItem('d365_app_updater_creds_temp'); // Clean up temp creds on error too
        logError('=== AUTH FAILED ===', e.message);
        hideLoading();
        
        // CRITICAL: Clear ALL saved state to prevent auto-login loop
        localStorage.removeItem('d365_app_updater_creds');
        
        // Clear ALL MSAL cache (localStorage AND sessionStorage)
        Object.keys(localStorage).forEach(key => {
            if (key.startsWith('msal.')) localStorage.removeItem(key);
        });
        Object.keys(sessionStorage).forEach(key => {
            if (key.startsWith('msal.')) sessionStorage.removeItem(key);
        });
        
        // Clean URL hash if present (prevents re-processing on next attempt)
        if (window.location.hash) {
            history.replaceState(null, '', window.location.pathname);
        }
        
        // Make sure auth section is visible
        document.getElementById('authSection').classList.remove('hidden');
        document.getElementById('appsSection').classList.add('hidden');
        
        // Clear the form's "remember me" checkbox
        const rememberMe = document.getElementById('rememberMe');
        if (rememberMe) rememberMe.checked = false;
        
        // Provide helpful error messages for common issues
        let errorMessage = e.message;
        const resetButton = `<br><br><button onclick="window.location.reload()" style="background:#0078d4;color:white;border:none;padding:10px 20px;border-radius:6px;cursor:pointer;font-weight:600;">🔄 Start Fresh</button>`;
        
        if (e.message.includes('AADSTS650057') || e.message.includes('Invalid resource') || e.message.includes('not listed in the requested permissions')) {
            errorMessage = `<strong>Missing API Permissions</strong><br><br>
Your Azure AD app registration is missing required permissions.<br><br>
<strong>To fix this:</strong><br>
1. Go to <a href="https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">Azure Portal → App Registrations</a><br>
2. Find and click on your app<br>
3. Go to <strong>API permissions</strong> → <strong>Add a permission</strong><br>
4. Add: <strong>Power Platform API</strong> → Delegated → <code>user_impersonation</code><br>
5. Add: <strong>Dynamics CRM</strong> → Delegated → <code>user_impersonation</code><br>
6. Click <strong>Grant admin consent</strong>${resetButton}<br><br>
<small style="color:#888">Error: ${e.message.substring(0, 150)}...</small>`;
        } else if (e.message.includes('AADSTS700016') || e.message.includes('not found in the directory')) {
            errorMessage = `<strong>Application Not Found</strong><br><br>
The Client ID does not exist in the specified tenant.<br>
Please verify your Tenant ID and Client ID are correct.${resetButton}`;
        } else if (e.message.includes('AADSTS50011') || e.message.includes('reply URL') || e.message.includes('redirect')) {
            errorMessage = `<strong>Invalid Redirect URI</strong><br><br>
The redirect URI is not configured in your app registration.<br><br>
Add this URI to your app's redirect URIs:<br>
<code>${window.location.origin + window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1)}</code>${resetButton}`;
        } else {
            errorMessage = `<strong>Authentication Failed</strong><br><br>${e.message}${resetButton}`;
        }
        
        showError(errorMessage);
        msalInstance = null;
        ppToken = null;
        environmentId = null;
    }
}

// Handle authentication
async function handleAuthentication(event) {
    event.preventDefault();
    
    logInfo('=== handleAuthentication START (user clicked Connect) ===');
    
    let orgUrlValue = document.getElementById('orgUrl').value.trim();
    const tenantId = ''; // Tenant ID removed from UI - will use 'organizations' endpoint
    const clientId = document.getElementById('clientId').value.trim();
    const rememberMe = document.getElementById('rememberMe').checked;
    
    logInfo('Form values', { orgUrl: orgUrlValue, clientId: clientId.substring(0,8) + '...', rememberMe });
    
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(clientId)) {
        logError('Invalid Client ID GUID format');
        showError('Invalid Client ID format. Client ID must be a valid GUID.');
        return;
    }
    
    // Normalize the org URL
    if (!orgUrlValue.startsWith('https://')) {
        orgUrlValue = 'https://' + orgUrlValue;
    }
    orgUrlValue = orgUrlValue.replace(/\/+$/, ''); // remove trailing slashes
    
    if (!orgUrlValue.includes('.dynamics.com')) {
        logError('Invalid Organization URL');
        showError('Invalid Organization URL. It should look like https://yourorg.crm.dynamics.com');
        return;
    }
    
    // ALWAYS save to sessionStorage for redirect flow (survives redirect, works with tracking prevention)
    const credsObj = { orgUrl: orgUrlValue, tenantId, clientId, rememberMe };
    sessionStorage.setItem('d365_app_updater_creds_temp', JSON.stringify(credsObj));
    logInfo('Credentials saved to sessionStorage for redirect');
    
    // Also try localStorage if user wants to remember (may fail with tracking prevention)
    if (rememberMe) {
        try {
            localStorage.setItem('d365_app_updater_creds', JSON.stringify({ orgUrl: orgUrlValue, tenantId, clientId }));
            logInfo('Credentials also saved to localStorage');
        } catch (e) {
            logWarn('Could not save to localStorage (tracking prevention?)', e.message);
        }
    }
    
    // Clear any previous auth step state to start fresh
    sessionStorage.removeItem('d365_auth_step');
    logInfo('Cleared previous auth step');
    
    try {
        showLoading('Authenticating...', 'Connecting to Microsoft');
        
        const msalConfig = createMsalConfig(tenantId, clientId);
        logDebug('MSAL config', { redirectUri: msalConfig.auth.redirectUri, authority: msalConfig.auth.authority });
        
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        logInfo('MSAL initialized');
        
        // Redirect to Microsoft login — the page will reload and handleRedirectResponse() will continue
        showLoading('Authenticating...', 'Redirecting to Microsoft sign-in...');
        
        // Save the pending auth state so we know to continue after redirect
        sessionStorage.setItem('d365_auth_step', 'login_redirect');
        logInfo('Set auth step to login_redirect, calling loginRedirect...');
        
        await msalInstance.loginRedirect({
            scopes: ['openid', 'profile']
        });
        
        // Execution stops here — browser navigates to Microsoft login
        return;
        
    } catch (error) {
        hideLoading();
        console.error('Auth error:', error);
        showError('Authentication failed: ' + error.message);
    }
}

// Resolve an Organization URL (e.g. https://orgname.crm.dynamics.com) to a Power Platform Environment ID
async function resolveOrgUrlToEnvironmentId(orgUrl) {
    const bapToken = await getBAPToken();
    
    // Normalize for comparison: lowercase, no trailing slash
    const normalizedInput = orgUrl.toLowerCase().replace(/\/+$/, '');
    
    // List all environments and find the one whose instanceUrl matches
    const response = await fetch('https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2021-04-01', {
        headers: { 'Authorization': `Bearer ${bapToken}` }
    });
    
    if (!response.ok) {
        console.error('Failed to list environments:', response.status);
        throw new Error('Failed to list environments. Make sure you have Power Platform admin access.');
    }
    
    const data = await response.json();
    const environments = data.value || [];
    console.log('Found', environments.length, 'environments, searching for URL:', orgUrl);
    
    // Cache all environments for the switcher
    allEnvironments = environments.filter(env => env.properties?.linkedEnvironmentMetadata?.instanceUrl).map(env => ({
        id: env.name,
        name: env.properties?.displayName || env.name,
        instanceUrl: (env.properties?.linkedEnvironmentMetadata?.instanceUrl || '').replace(/\/+$/, ''),
        type: env.properties?.environmentType || '',
    })).sort((a, b) => a.name.localeCompare(b.name));
    
    for (const env of environments) {
        const instanceUrl = env.properties?.linkedEnvironmentMetadata?.instanceUrl;
        const envName = env.properties?.displayName || env.name;
        
        if (instanceUrl) {
            const normalizedInstance = instanceUrl.toLowerCase().replace(/\/+$/, '');
            console.log(`  Environment: ${envName} (${env.name}), instanceUrl: ${instanceUrl}`);
            
            if (normalizedInstance === normalizedInput) {
                console.log('  ✓ Match found! Environment ID:', env.name);
                return env.name;
            }
        }
    }
    
    console.error('No environment found matching URL:', orgUrl);
    return null;
}

// Get environment name from BAP API (for display purposes)
async function getEnvironmentName() {
    // Use shared BAP token helper
    const bapToken = await getBAPToken();
    
    // Get environment details by ID
    const response = await fetch(`https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentId}?api-version=2021-04-01`, {
        headers: { 'Authorization': `Bearer ${bapToken}` }
    });
    
    if (response.ok) {
        const env = await response.json();
        const displayName = env.properties?.displayName || environmentId;
        document.getElementById('environmentName').textContent = displayName;
        console.log('Environment:', displayName);
    } else {
        // If we can't get the name, just use the ID
        document.getElementById('environmentName').textContent = environmentId;
        console.log('Could not get environment name, using ID');
    }
    
    // Render environment switcher
    renderEnvSwitcher();
}

// ── Environment Switcher ──────────────────────────────────────
function renderEnvSwitcher() {
    const list = document.getElementById('envList');
    if (!list || allEnvironments.length === 0) return;
    
    let html = '';
    for (const env of allEnvironments) {
        const isActive = env.id === environmentId;
        const shortUrl = env.instanceUrl.replace(/^https?:\/\//, '');
        html += '<div class="env-item' + (isActive ? ' active' : '') + '" onclick="switchEnvironment(\'' + env.id + '\')" title="' + escapeHtml(env.instanceUrl) + '">';
        html += '<div class="env-item-icon"><i class="fas fa-' + (isActive ? 'check' : 'globe') + '"></i></div>';
        html += '<div class="env-item-details">';
        html += '<div class="env-item-name">' + escapeHtml(env.name) + '</div>';
        html += '<div class="env-item-url">' + escapeHtml(shortUrl) + '</div>';
        html += '</div>';
        html += '</div>';
    }
    list.innerHTML = html;
}

function toggleEnvDropdown() {
    const dropdown = document.getElementById('envDropdown');
    const btn = document.getElementById('envSwitcherBtn');
    const isOpen = dropdown.classList.contains('show');
    if (isOpen) {
        closeEnvDropdown();
    } else {
        dropdown.classList.add('show');
        btn.classList.add('open');
        document.getElementById('envSearchInput').value = '';
        document.getElementById('envSearchInput').focus();
        filterEnvList();
    }
}

function closeEnvDropdown() {
    const dropdown = document.getElementById('envDropdown');
    const btn = document.getElementById('envSwitcherBtn');
    if (dropdown) dropdown.classList.remove('show');
    if (btn) btn.classList.remove('open');
}

function filterEnvList() {
    const search = (document.getElementById('envSearchInput').value || '').toLowerCase();
    const list = document.getElementById('envList');
    const filtered = allEnvironments.filter(env => {
        if (!search) return true;
        return env.name.toLowerCase().includes(search) || env.instanceUrl.toLowerCase().includes(search);
    });
    
    if (filtered.length === 0) {
        list.innerHTML = '<div class="env-empty"><i class="fas fa-search me-1"></i> No environments found</div>';
        return;
    }
    
    let html = '';
    for (const env of filtered) {
        const isActive = env.id === environmentId;
        const shortUrl = env.instanceUrl.replace(/^https?:\/\//, '');
        html += '<div class="env-item' + (isActive ? ' active' : '') + '" onclick="switchEnvironment(\'' + env.id + '\')" title="' + escapeHtml(env.instanceUrl) + '">';
        html += '<div class="env-item-icon"><i class="fas fa-' + (isActive ? 'check' : 'globe') + '"></i></div>';
        html += '<div class="env-item-details">';
        html += '<div class="env-item-name">' + escapeHtml(env.name) + '</div>';
        html += '<div class="env-item-url">' + escapeHtml(shortUrl) + '</div>';
        html += '</div>';
        html += '</div>';
    }
    list.innerHTML = html;
}

async function switchEnvironment(envId) {
    if (envId === environmentId) {
        closeEnvDropdown();
        return;
    }
    
    closeEnvDropdown();
    
    const env = allEnvironments.find(e => e.id === envId);
    if (!env) return;
    
    environmentId = envId;
    currentOrgUrl = env.instanceUrl;
    selectedApps.clear();
    
    // Update saved credentials with new org URL
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            creds.orgUrl = env.instanceUrl;
            localStorage.setItem('d365_app_updater_creds', JSON.stringify(creds));
        } catch (e) {}
    }
    
    document.getElementById('environmentName').textContent = env.name;
    renderEnvSwitcher();
    
    console.log('Switching to environment:', env.name, '(' + envId + ')');
    await loadApplications();
}

// Compare two version strings (e.g., "1.2.3.4" vs "1.2.3.5")
function compareVersions(v1, v2) {
    if (!v1 || !v2) return 0;
    const parts1 = v1.split('.').map(Number);
    const parts2 = v2.split('.').map(Number);
    for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
        const p1 = parts1[i] || 0;
        const p2 = parts2[i] || 0;
        if (p1 > p2) return 1;
        if (p1 < p2) return -1;
    }
    return 0;
}

// Helper: fetch all pages from a paginated Power Platform API endpoint
async function fetchAllPages(url, token) {
    let allItems = [];
    let nextUrl = url;
    let pageCount = 0;
    
    while (nextUrl) {
        pageCount++;
        console.log(`Fetching page ${pageCount}: ${nextUrl.substring(0, 120)}...`);
        
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 60000);
        
        const response = await fetch(nextUrl, {
            headers: { 'Authorization': `Bearer ${token}` },
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Error on page', pageCount, ':', response.status, errorText);
            throw new Error('Failed to fetch apps (page ' + pageCount + '): ' + response.status);
        }
        
        const data = await response.json();
        const items = data.value || [];
        allItems = allItems.concat(items);
        nextUrl = data['@odata.nextLink'] || null;
        
        console.log(`Page ${pageCount}: got ${items.length} items (total so far: ${allItems.length})`);
    }
    
    return allItems;
}

// Helper: get BAP token for admin API calls
async function getBAPToken() {
    const accounts = msalInstance.getAllAccounts();
    const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account: accounts[0] };
    try {
        const result = await msalInstance.acquireTokenSilent(bapRequest);
        return result.accessToken;
    } catch (e) {
        // Silent failed — redirect for consent, but track step to avoid loops
        const currentStep = sessionStorage.getItem('d365_auth_step');
        if (currentStep === 'acquiring_bap_runtime') {
            throw new Error('Failed to acquire BAP token. Please check your app registration permissions.');
        }
        sessionStorage.setItem('d365_auth_step', 'acquiring_bap_runtime');
        await msalInstance.acquireTokenRedirect(bapRequest);
        // Page will reload, throw to stop current execution
        throw new Error('Redirecting for BAP API consent...');
    }
}

// Load applications from Power Platform API
async function loadApplications() {
    showLoading('Loading applications...', 'Fetching from Power Platform');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>';
    
    try {
        const baseUrl = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages`;
        const apiVersion = 'api-version=2022-03-01-preview';
        
        // ── Step 1: Fetch INSTALLED apps explicitly ──────────────────
        showLoading('Loading applications...', 'Fetching installed apps...');
        const installedAppsRaw = await fetchAllPages(
            `${baseUrl}?appInstallState=Installed&${apiVersion}`, ppToken
        );
        console.log('Installed apps fetched:', installedAppsRaw.length);
        
        // ── Step 2: Fetch ALL catalog packages (includes newer versions) ──
        showLoading('Loading applications...', 'Fetching available catalog versions...');
        const allAppsRaw = await fetchAllPages(
            `${baseUrl}?${apiVersion}`, ppToken
        );
        console.log('All catalog packages fetched:', allAppsRaw.length);
        
        // ── Step 2b: Fetch NotInstalled packages specifically (update packages) ──
        showLoading('Loading applications...', 'Fetching update packages...');
        let notInstalledRaw = [];
        try {
            notInstalledRaw = await fetchAllPages(
                `${baseUrl}?appInstallState=NotInstalled&${apiVersion}`, ppToken
            );
            console.log('NotInstalled packages fetched:', notInstalledRaw.length);
        } catch (e) {
            console.warn('NotInstalled fetch failed (non-critical):', e.message);
        }
        
        // Merge all catalog sources
        const allCatalogEntries = [...allAppsRaw, ...notInstalledRaw];
        
        // Debug: log unique states and sample data
        const states = [...new Set(allCatalogEntries.map(a => a.state))];
        console.log('All states found in catalog:', states);
        if (installedAppsRaw.length > 0) {
            console.log('Sample installed app fields:', Object.keys(installedAppsRaw[0]));
            console.log('Sample installed app:', JSON.stringify(installedAppsRaw[0], null, 2));
        }
        
        // ── Step 3: Build version maps from ALL catalog entries ──────
        // Map by applicationId → keep highest version
        const catalogMapById = new Map();
        for (const app of allCatalogEntries) {
            if (!app.applicationId) continue;
            const existing = catalogMapById.get(app.applicationId);
            if (!existing || compareVersions(app.version, existing.version) > 0) {
                catalogMapById.set(app.applicationId, app);
            }
        }
        
        // Map by uniqueName base → keep highest version (fallback matching)
        const catalogByName = new Map();
        for (const app of allCatalogEntries) {
            if (!app.uniqueName) continue;
            const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
            const existing = catalogByName.get(baseName);
            if (!existing || compareVersions(app.version, existing.version) > 0) {
                catalogByName.set(baseName, app);
            }
        }
        
        // Map by display name → keep highest version
        const catalogByDisplayName = new Map();
        for (const app of allCatalogEntries) {
            const name = (app.localizedName || app.applicationName || '').toLowerCase();
            if (!name) continue;
            const existing = catalogByDisplayName.get(name);
            if (!existing || compareVersions(app.version, existing.version) > 0) {
                catalogByDisplayName.set(name, app);
            }
        }
        
        console.log('Catalog map by ID entries:', catalogMapById.size);
        console.log('Catalog map by name entries:', catalogByName.size);
        
        // ── Step 4: Detect updates for each installed app ────────────
        let updatesFound = 0;
        apps = installedAppsRaw.map(app => {
            let hasUpdate = false;
            let latestVersion = null;
            let catalogUniqueName = null;
            let spaOnly = false;
            
            // Skip update detection for apps that require Custom Install Experience (SPA)
            // These cannot be updated via the API — they must be updated through the Admin Center
            if (app.singlePageApplicationUrl) {
                spaOnly = true;
            }
            
            // Check 0: State-based detection — API may directly flag updates
            const stateLower = (app.state || '').toLowerCase();
            if (!spaOnly && (stateLower.includes('update') || stateLower === 'updateavailable' || stateLower === 'installedwithupdateavailable')) {
                hasUpdate = true;
                console.log(`  [by state="${app.state}"] ${app.localizedName || app.uniqueName}`);
            }
            
            // Check 1: Direct API fields that might indicate update availability
            if (!spaOnly && (app.updateAvailable || app.catalogVersion || app.availableVersion || app.latestVersion || app.newVersion || app.updateVersion)) {
                const directVersion = app.catalogVersion || app.availableVersion || app.latestVersion || app.newVersion || app.updateVersion;
                if (directVersion && compareVersions(directVersion, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = directVersion;
                    console.log(`  [direct field] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                }
                // updateAvailable might be a boolean
                if (app.updateAvailable === true && !latestVersion) {
                    hasUpdate = true;
                    console.log(`  [updateAvailable=true] ${app.localizedName || app.uniqueName}`);
                }
            }
            
            // Check 2: Compare with catalog entry by applicationId
            if (!spaOnly && !hasUpdate && app.applicationId) {
                const catalogEntry = catalogMapById.get(app.applicationId);
                if (catalogEntry && compareVersions(catalogEntry.version, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = catalogEntry.version;
                    catalogUniqueName = catalogEntry.uniqueName;
                    console.log(`  [by appId] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                }
            }
            
            // Check 3: Compare with catalog entry by uniqueName base
            if (!spaOnly && !hasUpdate && app.uniqueName) {
                const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                const byName = catalogByName.get(baseName);
                if (byName && compareVersions(byName.version, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = byName.version;
                    catalogUniqueName = byName.uniqueName;
                    console.log(`  [by name] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                }
            }
            
            // Check 4: Compare with catalog entry by localizedName / applicationName
            if (!spaOnly && !hasUpdate) {
                const appName = (app.localizedName || app.applicationName || '').toLowerCase();
                if (appName) {
                    for (const [, catApp] of catalogMapById) {
                        const catName = (catApp.localizedName || catApp.applicationName || '').toLowerCase();
                        if (catName === appName && compareVersions(catApp.version, app.version) > 0) {
                            hasUpdate = true;
                            latestVersion = catApp.version;
                            catalogUniqueName = catApp.uniqueName;
                            console.log(`  [by displayName] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                            break;
                        }
                    }
                }
            }
            
            if (hasUpdate) updatesFound++;
            if (spaOnly) {
                console.log(`  [skipped SPA] ${app.localizedName || app.uniqueName} — requires Admin Center`);
            }
            
            return {
                id: app.id,
                uniqueName: app.uniqueName,
                catalogUniqueName: catalogUniqueName || app.uniqueName,
                name: app.localizedName || app.applicationName || app.uniqueName || 'Unknown',
                version: app.version || 'Unknown',
                latestVersion: latestVersion,
                state: app.state || 'Installed',
                hasUpdate: hasUpdate,
                publisher: app.publisherName || 'Microsoft',
                description: app.applicationDescription || '',
                learnMoreUrl: app.learnMoreUrl || null,
                instancePackageId: app.instancePackageId,
                applicationId: app.applicationId,
                spaOnly: spaOnly
            };
        });
        
        console.log('Updates found from PP API:', updatesFound);
        
        // ── Step 5: ALWAYS check BAP Admin API for additional updates ──
        // The BAP API can detect updates that the PP API misses
        console.log('Checking BAP Admin API for additional updates...');
        showLoading('Loading applications...', 'Cross-checking updates via Admin API...');
        
        try {
            const bapToken = await getBAPToken();
            const bapUrl = `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentId}/applicationPackages?api-version=2021-04-01`;
            const bapApps = await fetchAllPages(bapUrl, bapToken);
            
            console.log('BAP API returned:', bapApps.length, 'packages');
            if (bapApps.length > 0) {
                console.log('BAP sample app fields:', Object.keys(bapApps[0]));
                // Log first 3 samples for debugging
                bapApps.slice(0, 3).forEach((a, i) => console.log(`BAP sample ${i}:`, JSON.stringify(a, null, 2)));
            }
            
            // Build BAP catalog map by applicationId (keep highest version)
            const bapCatalogMap = new Map();
            for (const bapApp of bapApps) {
                if (!bapApp.applicationId) continue;
                const existing = bapCatalogMap.get(bapApp.applicationId);
                if (!existing || compareVersions(bapApp.version, existing.version) > 0) {
                    bapCatalogMap.set(bapApp.applicationId, bapApp);
                }
            }
            
            // Also build by uniqueName base
            const bapByName = new Map();
            for (const bapApp of bapApps) {
                if (!bapApp.uniqueName) continue;
                const baseName = bapApp.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                const existing = bapByName.get(baseName);
                if (!existing || compareVersions(bapApp.version, existing.version) > 0) {
                    bapByName.set(baseName, bapApp);
                }
            }
            
            // Also build by display name
            const bapByDisplayName = new Map();
            for (const bapApp of bapApps) {
                const name = (bapApp.localizedName || bapApp.applicationName || '').toLowerCase();
                if (!name) continue;
                const existing = bapByDisplayName.get(name);
                if (!existing || compareVersions(bapApp.version, existing.version) > 0) {
                    bapByDisplayName.set(name, bapApp);
                }
            }
            
            // Check installed apps that DON'T already have an update detected
            for (const app of apps) {
                if (app.hasUpdate) continue;
                
                let found = false;
                
                // Check direct fields from BAP response for this app
                const bapInstalled = bapApps.find(b => 
                    (b.applicationId === app.applicationId) && 
                    (b.state === 'Installed' || b.instancePackageId)
                );
                if (bapInstalled) {
                    // State-based detection
                    const bapState = (bapInstalled.state || '').toLowerCase();
                    if (bapState.includes('update') || bapState === 'updateavailable') {
                        app.hasUpdate = true;
                        found = true;
                        console.log(`  [BAP state="${bapInstalled.state}"] ${app.name}`);
                    }
                    // Check if BAP provides update info directly
                    const directVer = bapInstalled.catalogVersion || bapInstalled.availableVersion || bapInstalled.latestVersion || bapInstalled.newVersion || bapInstalled.updateVersion;
                    if (directVer && compareVersions(directVer, app.version) > 0) {
                        app.hasUpdate = true;
                        app.latestVersion = directVer;
                        found = true;
                        console.log(`  [BAP direct] ${app.name}: ${app.version} → ${directVer}`);
                    }
                    if (bapInstalled.updateAvailable === true) {
                        app.hasUpdate = true;
                        found = true;
                        console.log(`  [BAP updateAvailable=true] ${app.name}`);
                    }
                }
                
                // Compare by applicationId
                if (!found && app.applicationId) {
                    const bapEntry = bapCatalogMap.get(app.applicationId);
                    if (bapEntry && compareVersions(bapEntry.version, app.version) > 0) {
                        app.hasUpdate = true;
                        app.latestVersion = bapEntry.version;
                        app.catalogUniqueName = bapEntry.uniqueName || app.uniqueName;
                        found = true;
                        console.log(`  [BAP by appId] ${app.name}: ${app.version} → ${bapEntry.version}`);
                    }
                }
                
                // Compare by uniqueName
                if (!found && app.uniqueName) {
                    const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                    const bapByNameEntry = bapByName.get(baseName);
                    if (bapByNameEntry && compareVersions(bapByNameEntry.version, app.version) > 0) {
                        app.hasUpdate = true;
                        app.latestVersion = bapByNameEntry.version;
                        app.catalogUniqueName = bapByNameEntry.uniqueName || app.uniqueName;
                        found = true;
                        console.log(`  [BAP by name] ${app.name}: ${app.version} → ${bapByNameEntry.version}`);
                    }
                }
                
                // Compare by display name
                if (!found) {
                    const appDisplayName = (app.name || '').toLowerCase();
                    if (appDisplayName) {
                        const bapByDN = bapByDisplayName.get(appDisplayName);
                        if (bapByDN && compareVersions(bapByDN.version, app.version) > 0) {
                            app.hasUpdate = true;
                            app.latestVersion = bapByDN.version;
                            app.catalogUniqueName = bapByDN.uniqueName || app.uniqueName;
                            found = true;
                            console.log(`  [BAP by displayName] ${app.name}: ${app.version} → ${bapByDN.version}`);
                        }
                    }
                }
                
                if (found) updatesFound++;
            }
            
            // ── Step 5b: Check for installed apps that BAP knows but PP API missed entirely ──
            const knownAppIds = new Set(apps.map(a => a.applicationId).filter(Boolean));
            const knownNames = new Set(apps.map(a => (a.name || '').toLowerCase()).filter(Boolean));
            
            for (const bapApp of bapApps) {
                const bapState = (bapApp.state || '').toLowerCase();
                const isInstalled = bapState === 'installed' || bapState.includes('update') || bapApp.instancePackageId;
                if (!isInstalled) continue;
                
                // Skip if we already know about this app
                if (bapApp.applicationId && knownAppIds.has(bapApp.applicationId)) continue;
                const bapName = (bapApp.localizedName || bapApp.applicationName || '').toLowerCase();
                if (bapName && knownNames.has(bapName)) continue;
                
                // This is an installed app the PP API missed
                let hasUpdate = false;
                let latestVersion = null;
                
                if (bapState.includes('update') || bapApp.updateAvailable === true) {
                    hasUpdate = true;
                }
                const directVer = bapApp.catalogVersion || bapApp.availableVersion || bapApp.latestVersion || bapApp.newVersion;
                if (directVer && compareVersions(directVer, bapApp.version) > 0) {
                    hasUpdate = true;
                    latestVersion = directVer;
                }
                // Check if BAP catalog has a higher version
                if (!hasUpdate && bapApp.applicationId) {
                    const bapCatEntry = bapCatalogMap.get(bapApp.applicationId);
                    if (bapCatEntry && compareVersions(bapCatEntry.version, bapApp.version) > 0) {
                        hasUpdate = true;
                        latestVersion = bapCatEntry.version;
                    }
                }
                
                if (hasUpdate) {
                    updatesFound++;
                    console.log(`  [BAP new app] ${bapApp.localizedName || bapApp.uniqueName}: ${bapApp.version} → ${latestVersion || 'update flagged'}`);
                }
                
                apps.push({
                    id: bapApp.id,
                    uniqueName: bapApp.uniqueName,
                    catalogUniqueName: bapApp.uniqueName,
                    name: bapApp.localizedName || bapApp.applicationName || bapApp.uniqueName || 'Unknown',
                    version: bapApp.version || 'Unknown',
                    latestVersion: latestVersion,
                    state: bapApp.state || 'Installed',
                    hasUpdate: hasUpdate,
                    publisher: bapApp.publisherName || 'Microsoft',
                    description: bapApp.applicationDescription || '',
                    learnMoreUrl: bapApp.learnMoreUrl || null,
                    instancePackageId: bapApp.instancePackageId,
                    applicationId: bapApp.applicationId
                });
            }
            
            console.log('Total updates found after BAP cross-check:', updatesFound);
        } catch (bapError) {
            console.warn('BAP API cross-check failed (non-critical):', bapError.message);
        }
        
        // ── Step 6: Sort — updates first, then alphabetically ────────
        apps.sort((a, b) => {
            if (a.hasUpdate && !b.hasUpdate) return -1;
            if (!a.hasUpdate && b.hasUpdate) return 1;
            return a.name.localeCompare(b.name);
        });
        
        // Store not-installed apps for browsing
        const knownInstalledAppIds = new Set(apps.map(a => a.applicationId).filter(Boolean));
        const notInstalledApps = [];
        for (const [appId, app] of catalogMapById) {
            if (!knownInstalledAppIds.has(appId) && app.state !== 'Installed' && !app.instancePackageId) {
                notInstalledApps.push({
                    id: app.id,
                    uniqueName: app.uniqueName,
                    catalogUniqueName: app.uniqueName,
                    name: app.localizedName || app.applicationName || app.uniqueName || 'Unknown',
                    version: app.version || 'Unknown',
                    latestVersion: null,
                    state: 'Available',
                    hasUpdate: false,
                    publisher: app.publisherName || 'Microsoft',
                    description: app.applicationDescription || '',
                    learnMoreUrl: app.learnMoreUrl || null,
                    instancePackageId: null,
                    applicationId: app.applicationId
                });
            }
        }
        notInstalledApps.sort((a, b) => a.name.localeCompare(b.name));
        window.availableApps = notInstalledApps;
        
        console.log('Final result:', apps.length, 'installed apps,', updatesFound, 'with updates');
        displayApplications();
        hideLoading();
        
    } catch (error) {
        hideLoading();
        console.error('Error loading applications:', error);
        
        let errorMsg = error.message;
        if (error.name === 'AbortError') {
            errorMsg = 'Request timed out. The Power Platform API took too long to respond. Please try again.';
        }
        
        const appsList = document.getElementById('appsList');
        appsList.innerHTML = '<div class="alert alert-danger"><i class="fas fa-exclamation-triangle me-2"></i>Failed to load applications: ' + errorMsg + '</div>';
    }
}

// Display applications
function displayApplications() {
    const appsList = document.getElementById('appsList');
    
    if (apps.length === 0) {
        appsList.innerHTML = '<div class="text-center py-5"><p class="text-muted">No applications found.</p></div>';
        return;
    }
    
    const appsWithUpdates = apps.filter(a => a.hasUpdate && !a.updateState);
    const appsUpdating = apps.filter(a => a.updateState === 'submitted' || a.updateState === 'updating');
    const appsFailed = apps.filter(a => a.updateState === 'failed');
    const installedApps = apps.filter(a => a.instancePackageId);
    const updateCount = appsWithUpdates.length;
    const updatingCount = appsUpdating.length;
    const failedCount = appsFailed.length;
    
    // Update summary text
    let summaryParts = [installedApps.length + ' apps installed'];
    if (updateCount > 0) {
        summaryParts.push('<span style="color: #28a745; font-weight: 600;">' + updateCount + ' update' + (updateCount !== 1 ? 's' : '') + ' available</span>');
    }
    if (updatingCount > 0) {
        summaryParts.push('<span style="color: #0d6efd; font-weight: 600;"><span class="spinner-updating"></span>' + updatingCount + ' updating</span>');
    }
    if (failedCount > 0) {
        summaryParts.push('<span style="color: #dc3545; font-weight: 600;">' + failedCount + ' failed</span>');
    }
    if (updateCount === 0 && updatingCount === 0 && failedCount === 0) {
        summaryParts.push('all up to date');
    }
    document.getElementById('appCountText').innerHTML = summaryParts.join(' &nbsp;|&nbsp; ');
    
    document.getElementById('updateAllBtn').disabled = updateCount === 0;
    
    // Update selected button visibility
    updateSelectedButton();
    
    // Show installed apps (include updating/failed states)
    const installedOrUpdatable = apps.filter(a => a.hasUpdate || a.instancePackageId || a.updateState);
    const appsToShow = installedOrUpdatable.length > 0 ? installedOrUpdatable : apps.slice(0, 50);
    
    // Sort: failed first, then updating, then updates available, then installed
    appsToShow.sort((a, b) => {
        const order = s => s === 'failed' ? 0 : (s === 'submitted' || s === 'updating') ? 1 : 2;
        const oa = order(a.updateState), ob = order(b.updateState);
        if (oa !== ob) return oa - ob;
        if (a.hasUpdate && !b.hasUpdate) return -1;
        if (!a.hasUpdate && b.hasUpdate) return 1;
        return a.name.localeCompare(b.name);
    });
    
    let html = '';
    
    // Status banners
    if (failedCount > 0) {
        html += '<div class="alert alert-danger mb-3" style="border-left: 4px solid #dc3545;">';
        html += '<div class="d-flex align-items-center">';
        html += '<i class="fas fa-exclamation-triangle fa-2x me-3 text-danger"></i>';
        html += '<div>';
        html += '<strong>' + failedCount + ' update' + (failedCount !== 1 ? 's' : '') + ' failed</strong><br>';
        html += '<small>Scroll down to see details. You can retry individual apps or check the Power Platform Admin Center.</small>';
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    if (updatingCount > 0) {
        html += '<div class="alert alert-info mb-3" style="border-left: 4px solid #0d6efd;">';
        html += '<div class="d-flex align-items-center">';
        html += '<i class="fas fa-sync-alt fa-spin fa-2x me-3 text-primary"></i>';
        html += '<div>';
        html += '<strong>' + updatingCount + ' update' + (updatingCount !== 1 ? 's' : '') + ' in progress</strong><br>';
        html += '<small>Updates are running in the background. Click <strong>"Refresh"</strong> to check current status.</small>';
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    if (updateCount > 0) {
        html += '<div class="alert alert-warning mb-3" style="border-left: 4px solid #ffc107;">';
        html += '<div class="d-flex align-items-center">';
        html += '<i class="fas fa-arrow-circle-up fa-2x me-3 text-warning"></i>';
        html += '<div>';
        html += '<strong>' + updateCount + ' update' + (updateCount !== 1 ? 's' : '') + ' available</strong><br>';
        html += '<small>Click <strong>"Update All Apps"</strong> to apply all updates, or update apps individually below.</small>';
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    if (updateCount === 0 && updatingCount === 0 && failedCount === 0) {
        html += '<div class="alert alert-success mb-3">';
        html += '<i class="fas fa-check-circle me-2"></i>';
        html += 'All installed applications are up to date.';
        html += '</div>';
    }
    
    for (const app of appsToShow) {
        let stateClass, stateIcon, stateText, cardClass, cardStyle;
        
        if (app.updateState === 'submitted' || app.updateState === 'updating') {
            stateClass = 'primary';
            stateIcon = '';
            stateText = 'Updating...';
            cardClass = 'app-card state-updating';
            cardStyle = '';
        } else if (app.updateState === 'failed') {
            stateClass = 'danger';
            stateIcon = 'exclamation-triangle';
            stateText = 'Failed';
            cardClass = 'app-card state-failed';
            cardStyle = '';
        } else if (app.hasUpdate) {
            stateClass = 'warning';
            stateIcon = 'arrow-circle-up';
            stateText = 'Update Available';
            cardClass = 'app-card';
            cardStyle = 'border-left: 4px solid #e67e22; background: #fef9f3;';
        } else {
            stateClass = 'secondary';
            stateIcon = 'check-circle';
            stateText = app.instancePackageId ? 'Installed' : 'Available';
            cardClass = 'app-card';
            cardStyle = '';
        }
        
        html += '<div class="' + cardClass + '" style="' + cardStyle + '">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-6">';
        // Show checkbox for apps with available updates or failed state
        const showCheckbox = app.hasUpdate || app.updateState === 'failed';
        const isChecked = selectedApps.has(app.uniqueName);
        if (showCheckbox) {
            html += '<div class="d-flex align-items-start">';
            html += '<input type="checkbox" class="app-select-cb" ' + (isChecked ? 'checked' : '') + ' onchange="toggleAppSelection(\'' + escapeHtml(app.uniqueName) + '\', this.checked)" title="Select for bulk update">';
            html += '<div>';
        }
        html += '<div class="app-name"><i class="fas fa-cube me-2"></i>' + escapeHtml(app.name) + '</div>';
        html += '<div class="app-version mt-2">';
        html += '<i class="fas fa-tag"></i> Version: <strong>' + escapeHtml(app.version) + '</strong>';
        if ((app.hasUpdate || app.updateState === 'submitted' || app.updateState === 'updating') && app.latestVersion) {
            html += ' <i class="fas fa-long-arrow-alt-right text-' + (app.updateState ? 'primary' : 'success') + ' mx-1"></i> <strong class="text-' + (app.updateState ? 'primary' : 'success') + '">' + escapeHtml(app.latestVersion) + '</strong>';
        }
        html += '</div>';
        html += '<div class="text-muted small mt-1"><i class="fas fa-building"></i> ' + escapeHtml(app.publisher) + '</div>';
        if (app.updateState === 'failed' && app.updateError) {
            html += '<div class="error-detail" title="' + escapeHtml(app.updateError) + '"><i class="fas fa-exclamation-circle me-1"></i>' + escapeHtml(app.updateError) + '</div>';
        }
        if (showCheckbox) {
            html += '</div></div>'; // close checkbox wrapper divs
        }
        html += '</div>';
        html += '<div class="col-md-3 text-center">';
        if (app.updateState === 'submitted' || app.updateState === 'updating') {
            html += '<span class="badge bg-primary"><span class="spinner-updating"></span> Updating...</span>';
        } else {
            html += '<span class="badge bg-' + stateClass + '">';
            if (stateIcon) html += '<i class="fas fa-' + stateIcon + '"></i> ';
            html += stateText + '</span>';
        }
        html += '</div>';
        html += '<div class="col-md-3 text-end">';
        if (app.updateState === 'submitted' || app.updateState === 'updating') {
            html += '<span class="text-primary"><i class="fas fa-sync-alt fa-spin"></i> In progress</span>';
        } else if (app.updateState === 'failed') {
            html += '<button class="btn btn-outline-danger btn-sm" onclick="updateSingleApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-redo"></i> Retry</button>';
        } else if (app.hasUpdate) {
            html += '<button class="btn btn-success btn-sm" onclick="updateSingleApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-download"></i> Update</button>';
        } else if (!app.instancePackageId) {
            html += '<button class="btn btn-primary btn-sm" onclick="installApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-plus"></i> Install</button>';
        } else {
            html += '<span class="text-success"><i class="fas fa-check-circle"></i> Up to date</span>';
        }
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    
    appsList.innerHTML = html;
}

// Update a single app
async function updateSingleApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    // If retrying a failed update, skip confirmation
    if (app.updateState !== 'failed') {
        if (!(await showModal({ title: 'Update App', message: 'Install update for "' + app.name + '"?\n\nCurrent: ' + app.version + '\nNew: ' + (app.latestVersion || 'latest'), type: 'update', okText: 'Update', okClass: 'btn-success-modal' }))) {
            return;
        }
    }
    
    // Mark as updating and refresh display immediately
    app.updateState = 'submitted';
    app.updateError = null;
    displayApplications();
    
    try {
        const installUniqueName = app.catalogUniqueName || app.uniqueName;
        const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;
        
        console.log('Installing update:', installUniqueName, 'for', app.name);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${ppToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            // Check if this is a SPA-only app that can't be updated via API
            if (response.status === 400 && errorText.includes('Custom Install Experience')) {
                throw new Error('This app requires manual update through the Power Platform Admin Center. It cannot be updated via API.');
            }
            throw new Error(response.status + ' - ' + errorText);
        }
        
        // Keep as 'submitted' — user can Refresh to check later
        app.updateState = 'submitted';
        app.hasUpdate = false; // Don't count it as a pending update anymore
        displayApplications();
        logUsage(1, 0, [app.name]);
        
    } catch (error) {
        console.error('Update error:', error);
        app.updateState = 'failed';
        app.updateError = error.message;
        app.hasUpdate = true; // Keep as updatable so retry is possible
        displayApplications();
        logUsage(0, 1, [app.name]);
    }
}

// Install an app
async function installApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    if (!(await showModal({ title: 'Install App', message: 'Install "' + app.name + '"?', type: 'info', okText: 'Install', okClass: 'btn-success-modal' }))) {
        return;
    }
    
    app.updateState = 'submitted';
    app.updateError = null;
    displayApplications();
    
    try {
        const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2022-03-01-preview`;
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${ppToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(response.status + ' - ' + errorText);
        }
        
        app.updateState = 'submitted';
        displayApplications();
        
    } catch (error) {
        console.error('Install error:', error);
        app.updateState = 'failed';
        app.updateError = error.message;
        displayApplications();
    }
}

// Update an installed app
async function reinstallApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    if (app.updateState !== 'failed') {
        if (!(await showModal({ title: 'Update App', message: 'Update "' + app.name + '"?\n\nCurrent version: ' + app.version, type: 'update', okText: 'Update', okClass: 'btn-success-modal' }))) {
            return;
        }
    }
    
    app.updateState = 'submitted';
    app.updateError = null;
    displayApplications();
    
    try {
        const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2022-03-01-preview`;
        
        console.log('Updating:', app.name);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${ppToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(response.status + ' - ' + errorText);
        }
        
        app.updateState = 'submitted';
        app.hasUpdate = false;
        displayApplications();
        
    } catch (error) {
        console.error('Update error:', error);
        app.updateState = 'failed';
        app.updateError = error.message;
        app.hasUpdate = true;
        displayApplications();
    }
}

// Update all apps
async function updateAllApps() {
    const appsToUpdate = apps.filter(a => a.hasUpdate && a.updateState !== 'submitted' && a.updateState !== 'updating');
    
    if (appsToUpdate.length === 0) {
        await showAlert('No Updates', 'No updates available.', 'info');
        return;
    }
    
    if (!(await showUpdateConfirm(appsToUpdate))) {
        return;
    }
    
    let successCount = 0;
    let failCount = 0;
    
    // Mark all as updating immediately
    for (const app of appsToUpdate) {
        app.updateState = 'submitted';
        app.updateError = null;
    }
    displayApplications();
    
    showLoading('Installing updates...', '0 of ' + appsToUpdate.length);
    
    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + appsToUpdate.length + ': ' + app.name;
        
        try {
            const installUniqueName = app.catalogUniqueName || app.uniqueName;
            const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;
            
            console.log('Updating:', app.name, 'using package:', installUniqueName);
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${ppToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.ok) {
                successCount++;
                app.updateState = 'submitted';
                app.hasUpdate = false;
            } else {
                failCount++;
                const errorText = await response.text();
                app.updateState = 'failed';
                app.updateError = response.status + ' - ' + errorText;
                app.hasUpdate = true;
                console.error('Failed to update ' + app.name + ':', response.status);
            }
        } catch (error) {
            failCount++;
            app.updateState = 'failed';
            app.updateError = error.message;
            app.hasUpdate = true;
            console.error('Error updating ' + app.name + ':', error);
        }
        
        // Small delay between requests
        await new Promise(r => setTimeout(r, 1000));
    }
    
    hideLoading();
    displayApplications();
    
    // Log usage to Supabase
    logUsage(successCount, failCount, appsToUpdate.map(a => a.name));
    
    if (failCount === 0) {
        await showAlert('Updates Started', 'All ' + successCount + ' updates submitted successfully! Updates are running in the background and may take several minutes. Click "Refresh" to check progress.', 'success');
    } else {
        await showAlert('Updates Submitted', successCount + ' update' + (successCount !== 1 ? 's' : '') + ' submitted successfully. ' + failCount + ' failed — see details below. You can retry failed updates individually.', 'warning');
    }
}

// Update all installed apps
async function reinstallAllApps() {
    const appsToUpdate = apps.filter(a => a.hasUpdate && a.updateState !== 'submitted' && a.updateState !== 'updating');
    
    if (appsToUpdate.length === 0) {
        await showAlert('All Up to Date', 'No updates available. All apps are up to date.', 'success');
        return;
    }
    
    if (!(await showUpdateConfirm(appsToUpdate))) {
        return;
    }
    
    let successCount = 0;
    let failCount = 0;
    
    // Mark all as updating immediately
    for (const app of appsToUpdate) {
        app.updateState = 'submitted';
        app.updateError = null;
    }
    displayApplications();
    
    showLoading('Updating apps...', '0 of ' + appsToUpdate.length);
    
    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + appsToUpdate.length + ': ' + app.name;
        
        try {
            const installUniqueName = app.catalogUniqueName || app.uniqueName;
            const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;
            
            console.log('Updating:', app.name, 'using package:', installUniqueName);
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${ppToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.ok) {
                successCount++;
                app.updateState = 'submitted';
                app.hasUpdate = false;
            } else {
                failCount++;
                const errorText = await response.text();
                app.updateState = 'failed';
                app.updateError = response.status + ' - ' + errorText;
                app.hasUpdate = true;
                console.error('Failed to update ' + app.name + ':', response.status);
            }
        } catch (error) {
            failCount++;
            app.updateState = 'failed';
            app.updateError = error.message;
            app.hasUpdate = true;
            console.error('Error updating ' + app.name + ':', error);
        }
        
        // Small delay between requests to avoid rate limiting
        await new Promise(r => setTimeout(r, 1500));
    }
    
    hideLoading();
    displayApplications();
    
    // Log usage to Supabase
    logUsage(successCount, failCount, appsToUpdate.map(a => a.name));
    
    if (failCount === 0) {
        await showAlert('Updates Submitted', 'All ' + successCount + ' update requests submitted! Updates are running in the background. Click "Refresh" to check progress.', 'success');
    } else {
        await showAlert('Updates Submitted', successCount + ' update' + (successCount !== 1 ? 's' : '') + ' submitted. ' + failCount + ' failed — see details below. You can retry failed updates individually.', 'warning');
    }
}

// Toggle individual app selection for multi-select
function toggleAppSelection(uniqueName, checked) {
    if (checked) {
        selectedApps.add(uniqueName);
    } else {
        selectedApps.delete(uniqueName);
    }
    updateSelectedButton();
}

// Show/hide the "Update Selected" button and update count
function updateSelectedButton() {
    const btn = document.getElementById('updateSelectedBtn');
    const countSpan = document.getElementById('selectedCount');
    if (!btn || !countSpan) return;
    const count = selectedApps.size;
    countSpan.textContent = count;
    if (count > 0) {
        btn.classList.remove('d-none');
    } else {
        btn.classList.add('d-none');
    }
}

// Update only the selected apps
async function updateSelectedApps() {
    if (selectedApps.size === 0) return;

    const appsToUpdate = apps.filter(a =>
        selectedApps.has(a.uniqueName) &&
        (a.hasUpdate || a.updateState === 'failed') &&
        a.updateState !== 'submitted' &&
        a.updateState !== 'updating'
    );

    if (appsToUpdate.length === 0) {
        await showAlert('Nothing to Update', 'The selected apps have no pending updates.', 'info');
        return;
    }

    if (!(await showUpdateConfirm(appsToUpdate))) {
        return;
    }

    let successCount = 0;
    let failCount = 0;

    // Mark all as updating immediately
    for (const app of appsToUpdate) {
        app.updateState = 'submitted';
        app.updateError = null;
    }
    displayApplications();

    showLoading('Updating selected apps...', '0 of ' + appsToUpdate.length);

    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + appsToUpdate.length + ': ' + app.name;

        try {
            const installUniqueName = app.catalogUniqueName || app.uniqueName;
            const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;

            console.log('Updating:', app.name, 'using package:', installUniqueName);

            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${ppToken}`,
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                successCount++;
                app.updateState = 'submitted';
                app.hasUpdate = false;
            } else {
                failCount++;
                const errorText = await response.text();
                app.updateState = 'failed';
                app.updateError = response.status + ' - ' + errorText;
                app.hasUpdate = true;
                console.error('Failed to update ' + app.name + ':', response.status);
            }
        } catch (error) {
            failCount++;
            app.updateState = 'failed';
            app.updateError = error.message;
            app.hasUpdate = true;
            console.error('Error updating ' + app.name + ':', error);
        }

        await new Promise(r => setTimeout(r, 1500));
    }

    hideLoading();
    selectedApps.clear();
    displayApplications();

    // Log usage to Supabase
    logUsage(successCount, failCount, appsToUpdate.map(a => a.name));

    if (failCount === 0) {
        await showAlert('Updates Submitted', 'All ' + successCount + ' selected update(s) submitted! Updates are running in the background. Click "Refresh" to check progress.', 'success');
    } else {
        await showAlert('Updates Submitted', successCount + ' submitted, ' + failCount + ' failed — see details below.', 'warning');
    }
}

// Logout
function handleLogout() {
    showModal({ title: 'Logout', message: 'Are you sure you want to logout?', type: 'warning', okText: 'Logout', okClass: 'btn-danger-modal' }).then(confirmed => {
        if (!confirmed) return;
        accessToken = null;
        ppToken = null;
        environmentId = null;
        apps = [];
        
        if (msalInstance) {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                // Clear MSAL cache so auto-login won't fire again
                msalInstance.logoutRedirect({ account: accounts[0] }).catch(() => {});
            } else {
                // No accounts but clear cache anyway
                msalInstance.clearCache().catch(() => {});
            }
        }
        
        document.getElementById('appsSection').classList.add('hidden');
        document.getElementById('authSection').classList.remove('hidden');
    });
}

// ── Usage Tracking (Supabase) ─────────────────────────────────
function getSupabaseConfig() {
    return { url: SUPABASE_URL, key: SUPABASE_KEY };
}

function getCurrentUserEmail() {
    if (!msalInstance) return null;
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        return accounts[0].username || accounts[0].name || null;
    }
    return null;
}

function getCurrentClientId() {
    // Get the client ID from stored credentials (used at login)
    const savedCreds = localStorage.getItem('d365_app_updater_creds') || 
                       sessionStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            return creds.clientId || '';
        } catch (e) {
            console.warn('Could not parse saved credentials');
        }
    }
    return '';
}

async function logUsage(successCount, failCount, appNames) {
    const cfg = getSupabaseConfig();
    if (!cfg) {
        console.log('Usage tracking: Supabase not configured, skipping.');
        return;
    }

    const record = {
        timestamp: new Date().toISOString(),
        user_email: getCurrentUserEmail() || 'unknown',
        org_url: currentOrgUrl || '',
        environment_id: environmentId || '',
        success_count: successCount || 0,
        fail_count: failCount || 0,
        total_apps: (successCount || 0) + (failCount || 0),
        app_names: (appNames || []).join(', ')
    };

    try {
        const resp = await fetch(`${cfg.url}/rest/v1/usage_logs`, {
            method: 'POST',
            headers: {
                'apikey': cfg.key,
                'Authorization': `Bearer ${cfg.key}`,
                'Content-Type': 'application/json',
                'Prefer': 'return=minimal'
            },
            body: JSON.stringify(record)
        });

        if (resp.ok) {
            console.log('Usage logged successfully:', record);
        } else {
            console.warn('Usage logging failed:', resp.status, await resp.text());
        }
    } catch (error) {
        console.warn('Usage logging error (non-critical):', error.message);
    }
}

// ── Auto-Update Scheduling (Supabase) ─────────────────────────────────
let scheduleLoaded = false;

function toggleScheduleDetails() {
    const details = document.getElementById('scheduleDetails');
    const enabled = document.getElementById('scheduleEnabled').checked;
    if (enabled) {
        details.style.display = details.style.display === 'none' ? 'block' : 'none';
    }
}

function handleScheduleToggle() {
    const enabled = document.getElementById('scheduleEnabled').checked;
    const details = document.getElementById('scheduleDetails');
    details.style.display = enabled ? 'block' : 'none';
    
    if (!enabled) {
        // Disable schedule in Supabase
        disableSchedule();
    }
}

function toggleSecretVisibility() {
    const secretInput = document.getElementById('scheduleClientSecret');
    const icon = document.getElementById('secretToggleIcon');
    if (secretInput.type === 'password') {
        secretInput.type = 'text';
        icon.className = 'fas fa-eye-slash';
    } else {
        secretInput.type = 'password';
        icon.className = 'fas fa-eye';
    }
}

function showCredentialsHelp() {
    showModal({
        title: 'How to Create an App Registration',
        message: `<div style="text-align: left; font-size: 0.9rem;">
<p><strong>1. Go to Azure Portal</strong><br>
Navigate to <a href="https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank">Azure AD App Registrations</a></p>

<p><strong>2. Create New Registration</strong><br>
- Click "New registration"<br>
- Name: "D365 App Updater"<br>
- Account type: "Single tenant"<br>
- Click "Register"</p>

<p><strong>3. Copy the Client ID</strong><br>
Copy the "Application (client) ID" from the Overview page</p>

<p><strong>4. Create a Client Secret</strong><br>
- Go to "Certificates & secrets"<br>
- Click "New client secret"<br>
- Add a description and expiry<br>
- Copy the secret value immediately</p>

<p><strong>5. Add API Permissions</strong><br>
- Go to "API permissions"<br>
- Add "Dynamics CRM" → "user_impersonation"<br>
- Click "Grant admin consent"</p>

<p><strong>6. Create Application User in Power Platform</strong><br>
- Go to <a href="https://admin.powerplatform.microsoft.com" target="_blank">Power Platform Admin Center</a><br>
- Select your environment → Settings → Users<br>
- Create Application User with the Client ID<br>
- Assign "System Administrator" role</p>
</div>`,
        type: 'info',
        confirmOnly: true
    });
}

async function loadSchedule() {
    if (scheduleLoaded) return;
    
    const cfg = getSupabaseConfig();
    if (!cfg) return;
    
    const userEmail = getCurrentUserEmail();
    const envId = environmentId || '';
    
    if (!userEmail || !envId) return;
    
    try {
        const resp = await fetch(
            `${cfg.url}/rest/v1/update_schedules?user_email=eq.${encodeURIComponent(userEmail)}&environment_id=eq.${encodeURIComponent(envId)}&select=*`,
            {
                headers: {
                    'apikey': cfg.key,
                    'Authorization': `Bearer ${cfg.key}`
                }
            }
        );
        
        if (resp.ok) {
            const schedules = await resp.json();
            if (schedules.length > 0) {
                const schedule = schedules[0];
                document.getElementById('scheduleEnabled').checked = schedule.enabled;
                document.getElementById('scheduleDay').value = schedule.day_of_week;
                document.getElementById('scheduleTime').value = schedule.time_utc;
                document.getElementById('scheduleTimezone').value = schedule.timezone || 'UTC';
                // Don't auto-fill secret for security, just indicate it's set
                if (schedule.client_secret) {
                    document.getElementById('scheduleClientSecret').placeholder = '••••••••••••••••';
                }
                document.getElementById('scheduleDetails').style.display = schedule.enabled ? 'block' : 'none';
                updateScheduleStatus(schedule);
            }
        }
        scheduleLoaded = true;
    } catch (error) {
        console.warn('Failed to load schedule:', error.message);
    }
}

function updateScheduleStatus(schedule) {
    const statusEl = document.getElementById('scheduleStatus');
    if (!schedule || !schedule.enabled) {
        statusEl.innerHTML = '<i class="fas fa-info-circle"></i> Schedule not configured';
        statusEl.className = 'schedule-status';
        return;
    }
    
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[schedule.day_of_week];
    const timeDisplay = formatTimeDisplay(schedule.time_utc);
    const lastRun = schedule.last_run_at ? new Date(schedule.last_run_at).toLocaleString() : 'Never';
    
    statusEl.innerHTML = `<i class="fas fa-check-circle"></i> Scheduled: Every <strong>${dayName}</strong> at <strong>${timeDisplay}</strong> (${schedule.timezone || 'UTC'})<br>
        <small>Last run: ${lastRun}</small>`;
    statusEl.className = 'schedule-status active';
}

function formatTimeDisplay(time24) {
    const [hours, minutes] = time24.split(':');
    const h = parseInt(hours, 10);
    const ampm = h >= 12 ? 'PM' : 'AM';
    const h12 = h % 12 || 12;
    return `${h12}:${minutes} ${ampm}`;
}

async function saveSchedule() {
    const cfg = getSupabaseConfig();
    if (!cfg) {
        showError('Scheduling requires Supabase configuration.');
        return;
    }
    
    const userEmail = getCurrentUserEmail();
    const envId = environmentId || '';
    const orgUrl = currentOrgUrl || '';
    
    if (!userEmail || !envId) {
        showError('Please connect to an environment first.');
        return;
    }
    
    const saveBtn = document.getElementById('scheduleSaveBtn');
    const originalText = saveBtn.innerHTML;
    saveBtn.disabled = true;
    saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
    
    const schedule = {
        user_email: userEmail,
        environment_id: envId,
        org_url: orgUrl,
        enabled: document.getElementById('scheduleEnabled').checked,
        day_of_week: parseInt(document.getElementById('scheduleDay').value, 10),
        time_utc: document.getElementById('scheduleTime').value,
        timezone: document.getElementById('scheduleTimezone').value,
        client_id: getCurrentClientId(),
        client_secret: document.getElementById('scheduleClientSecret').value.trim(),
        tenant_id: tenantId || '',
        updated_at: new Date().toISOString()
    };
    
    // Validate credentials if scheduling is enabled
    if (schedule.enabled && !schedule.client_secret) {
        saveBtn.disabled = false;
        saveBtn.innerHTML = originalText;
        showError('Client Secret is required for scheduled updates.');
        return;
    }
    
    try {
        // Upsert: try to update first, then insert if not exists
        const checkResp = await fetch(
            `${cfg.url}/rest/v1/update_schedules?user_email=eq.${encodeURIComponent(userEmail)}&environment_id=eq.${encodeURIComponent(envId)}&select=id`,
            {
                headers: {
                    'apikey': cfg.key,
                    'Authorization': `Bearer ${cfg.key}`
                }
            }
        );
        
        const existing = await checkResp.json();
        let resp;
        
        if (existing.length > 0) {
            // Update
            resp = await fetch(
                `${cfg.url}/rest/v1/update_schedules?id=eq.${existing[0].id}`,
                {
                    method: 'PATCH',
                    headers: {
                        'apikey': cfg.key,
                        'Authorization': `Bearer ${cfg.key}`,
                        'Content-Type': 'application/json',
                        'Prefer': 'return=representation'
                    },
                    body: JSON.stringify(schedule)
                }
            );
        } else {
            // Insert
            schedule.created_at = new Date().toISOString();
            resp = await fetch(
                `${cfg.url}/rest/v1/update_schedules`,
                {
                    method: 'POST',
                    headers: {
                        'apikey': cfg.key,
                        'Authorization': `Bearer ${cfg.key}`,
                        'Content-Type': 'application/json',
                        'Prefer': 'return=representation'
                    },
                    body: JSON.stringify(schedule)
                }
            );
        }
        
        if (resp.ok) {
            const saved = await resp.json();
            updateScheduleStatus(Array.isArray(saved) ? saved[0] : saved);
            
            // If scheduling is enabled, set up the app registration automatically
            if (schedule.enabled && schedule.client_id) {
                saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Setting up permissions...';
                
                const setupResult = await setupAppRegistration(schedule.client_id);
                
                if (setupResult.success) {
                    let message = 'Your auto-update schedule has been saved.';
                    if (setupResult.permissionsAdded) {
                        message += '<br><br>✅ Dynamics CRM permission added to your app registration.';
                    }
                    if (setupResult.appUserCreated) {
                        message += '<br>✅ Application user created in Dataverse.';
                    }
                    message += '<br><br>Updates will run automatically at the scheduled time.';
                    
                    showModal({
                        title: 'Schedule Saved',
                        message: message,
                        type: 'success',
                        confirmOnly: true
                    });
                } else {
                    showModal({
                        title: 'Schedule Saved (Manual Setup Needed)',
                        message: `Your schedule has been saved, but automatic setup failed:<br><br>
                            <strong>${setupResult.error}</strong><br><br>
                            Please manually:<br>
                            1. Add "Dynamics CRM → user_impersonation" permission to your app<br>
                            2. Grant admin consent<br>
                            3. Create an Application User in Power Platform Admin Center`,
                        type: 'warning',
                        confirmOnly: true
                    });
                }
            } else {
                showModal({
                    title: 'Schedule Saved',
                    message: 'Your auto-update schedule has been saved. Updates will run automatically at the scheduled time.',
                    type: 'success',
                    confirmOnly: true
                });
            }
        } else {
            const errorText = await resp.text();
            console.error('Failed to save schedule:', errorText);
            showError('Failed to save schedule. Please try again.');
        }
    } catch (error) {
        console.error('Schedule save error:', error);
        showError('Failed to save schedule: ' + error.message);
    } finally {
        saveBtn.disabled = false;
        saveBtn.innerHTML = originalText;
    }
}

// ── Automatic App Registration Setup ─────────────────────────────────
// Dynamics CRM API Resource ID (constant across all tenants)
const DYNAMICS_CRM_RESOURCE_ID = '00000007-0000-0000-c000-000000000000';
// user_impersonation scope ID for Dynamics CRM
const USER_IMPERSONATION_SCOPE_ID = '78ce3f0f-a1ce-49c2-8cde-64b5c0896db4';

async function setupAppRegistration(clientId) {
    // This function configures the user's app registration with required permissions
    // and creates the Application User in Dataverse
    
    if (!msalInstance) {
        console.warn('MSAL not initialized');
        return { success: false, error: 'Not authenticated' };
    }
    
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        return { success: false, error: 'No authenticated account' };
    }
    
    try {
        // Get Graph API token
        const graphToken = await msalInstance.acquireTokenSilent({
            scopes: ['https://graph.microsoft.com/Application.ReadWrite.All'],
            account: accounts[0]
        }).catch(async () => {
            // Fallback to popup if silent fails
            return await msalInstance.acquireTokenPopup({
                scopes: ['https://graph.microsoft.com/Application.ReadWrite.All'],
                account: accounts[0]
            });
        });
        
        if (!graphToken || !graphToken.accessToken) {
            return { success: false, error: 'Could not get Graph API permission. Please grant admin consent.' };
        }
        
        console.log('Got Graph API token, configuring app registration...');
        
        // Step 1: Find the app registration by client ID
        const appResp = await fetch(
            `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${clientId}'`,
            {
                headers: {
                    'Authorization': `Bearer ${graphToken.accessToken}`
                }
            }
        );
        
        if (!appResp.ok) {
            const errText = await appResp.text();
            console.error('Failed to find app registration:', errText);
            return { success: false, error: 'Could not find app registration. Make sure you have Application.ReadWrite.All permission.' };
        }
        
        const appData = await appResp.json();
        if (!appData.value || appData.value.length === 0) {
            return { success: false, error: `App registration with ID ${clientId} not found in your tenant.` };
        }
        
        const app = appData.value[0];
        const appObjectId = app.id;
        console.log('Found app registration:', app.displayName, 'Object ID:', appObjectId);
        
        // Step 2: Check if Dynamics CRM permission already exists
        const existingPermissions = app.requiredResourceAccess || [];
        const hasDynamicsCRM = existingPermissions.some(
            ra => ra.resourceAppId === DYNAMICS_CRM_RESOURCE_ID &&
                  ra.resourceAccess.some(a => a.id === USER_IMPERSONATION_SCOPE_ID)
        );
        
        if (!hasDynamicsCRM) {
            console.log('Adding Dynamics CRM user_impersonation permission...');
            
            // Add Dynamics CRM permission
            const newPermissions = [...existingPermissions];
            const dynamicsCrmEntry = newPermissions.find(ra => ra.resourceAppId === DYNAMICS_CRM_RESOURCE_ID);
            
            if (dynamicsCrmEntry) {
                // Add scope to existing entry
                if (!dynamicsCrmEntry.resourceAccess.some(a => a.id === USER_IMPERSONATION_SCOPE_ID)) {
                    dynamicsCrmEntry.resourceAccess.push({
                        id: USER_IMPERSONATION_SCOPE_ID,
                        type: 'Scope'
                    });
                }
            } else {
                // Add new entry for Dynamics CRM
                newPermissions.push({
                    resourceAppId: DYNAMICS_CRM_RESOURCE_ID,
                    resourceAccess: [{
                        id: USER_IMPERSONATION_SCOPE_ID,
                        type: 'Scope'
                    }]
                });
            }
            
            // Update the app registration
            const updateResp = await fetch(
                `https://graph.microsoft.com/v1.0/applications/${appObjectId}`,
                {
                    method: 'PATCH',
                    headers: {
                        'Authorization': `Bearer ${graphToken.accessToken}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        requiredResourceAccess: newPermissions
                    })
                }
            );
            
            if (!updateResp.ok) {
                const errText = await updateResp.text();
                console.error('Failed to update app permissions:', errText);
                return { success: false, error: 'Could not add Dynamics CRM permission. You may need to add it manually.' };
            }
            
            console.log('✅ Dynamics CRM permission added to app registration');
        } else {
            console.log('✅ Dynamics CRM permission already exists');
        }
        
        // Step 3: Grant admin consent (requires Directory.ReadWrite.All or admin privileges)
        // This creates a service principal if it doesn't exist and grants consent
        try {
            await grantAdminConsent(graphToken.accessToken, clientId);
        } catch (consentError) {
            console.warn('Admin consent may need to be granted manually:', consentError.message);
        }
        
        // Step 4: Create Application User in Dataverse
        const appUserResult = await createApplicationUser(clientId);
        
        return { 
            success: true, 
            permissionsAdded: !hasDynamicsCRM,
            appUserCreated: appUserResult.success,
            appUserMessage: appUserResult.message
        };
        
    } catch (error) {
        console.error('Setup error:', error);
        return { success: false, error: error.message };
    }
}

async function grantAdminConsent(graphToken, clientId) {
    // First ensure the service principal exists
    let spResp = await fetch(
        `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${clientId}'`,
        {
            headers: { 'Authorization': `Bearer ${graphToken}` }
        }
    );
    
    let spData = await spResp.json();
    let servicePrincipalId;
    
    if (!spData.value || spData.value.length === 0) {
        // Create service principal
        const createSpResp = await fetch(
            'https://graph.microsoft.com/v1.0/servicePrincipals',
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${graphToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ appId: clientId })
            }
        );
        
        if (createSpResp.ok) {
            const newSp = await createSpResp.json();
            servicePrincipalId = newSp.id;
            console.log('✅ Service principal created:', servicePrincipalId);
        } else {
            throw new Error('Could not create service principal');
        }
    } else {
        servicePrincipalId = spData.value[0].id;
    }
    
    // Get Dynamics CRM service principal
    const crmSpResp = await fetch(
        `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${DYNAMICS_CRM_RESOURCE_ID}'`,
        {
            headers: { 'Authorization': `Bearer ${graphToken}` }
        }
    );
    
    const crmSpData = await crmSpResp.json();
    if (!crmSpData.value || crmSpData.value.length === 0) {
        console.warn('Dynamics CRM service principal not found - consent may need to be granted manually');
        return;
    }
    
    const crmServicePrincipalId = crmSpData.value[0].id;
    
    // Grant oauth2PermissionGrant (delegated permission consent)
    const grantResp = await fetch(
        'https://graph.microsoft.com/v1.0/oauth2PermissionGrants',
        {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${graphToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                clientId: servicePrincipalId,
                consentType: 'AllPrincipals',
                resourceId: crmServicePrincipalId,
                scope: 'user_impersonation'
            })
        }
    );
    
    if (grantResp.ok || grantResp.status === 409) { // 409 = already exists
        console.log('✅ Admin consent granted for Dynamics CRM');
    } else {
        const errText = await grantResp.text();
        console.warn('Could not grant admin consent:', errText);
    }
}

async function createApplicationUser(clientId) {
    if (!currentOrgUrl || !msalInstance) {
        return { success: false, message: 'Not connected to environment' };
    }
    
    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            return { success: false, message: 'No authenticated account' };
        }
        
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: [`${currentOrgUrl}/.default`],
            account: accounts[0]
        });
        
        const accessToken = tokenResponse.accessToken;
        
        // Check if application user already exists
        const checkUrl = `${currentOrgUrl}/api/data/v9.2/systemusers?$filter=applicationid eq ${clientId}&$select=systemuserid,fullname`;
        const checkResp = await fetch(checkUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json'
            }
        });
        
        if (checkResp.ok) {
            const checkData = await checkResp.json();
            if (checkData.value && checkData.value.length > 0) {
                console.log('✅ Application user already exists:', checkData.value[0].fullname);
                return { success: true, message: 'Application user already exists' };
            }
        }
        
        // Get root business unit
        const buResp = await fetch(
            `${currentOrgUrl}/api/data/v9.2/businessunits?$filter=parentbusinessunitid eq null&$select=businessunitid`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json'
                }
            }
        );
        
        if (!buResp.ok) {
            return { success: false, message: 'Could not get business unit' };
        }
        
        const buData = await buResp.json();
        if (!buData.value || buData.value.length === 0) {
            return { success: false, message: 'No root business unit found' };
        }
        
        const businessUnitId = buData.value[0].businessunitid;
        
        // Create the application user
        const createResp = await fetch(`${currentOrgUrl}/api/data/v9.2/systemusers`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                'applicationid': clientId,
                'fullname': 'D365 App Updater Scheduler',
                'internalemailaddress': `app-updater-${clientId.substring(0, 8)}@automation.local`,
                'businessunitid@odata.bind': `/businessunits(${businessUnitId})`,
                'accessmode': 4 // Non-interactive (application user)
            })
        });
        
        if (createResp.ok || createResp.status === 204) {
            console.log('✅ Application user created successfully');
            
            // Get the created user ID and assign System Administrator role
            const userUrl = createResp.headers.get('OData-EntityId');
            if (userUrl) {
                const userIdMatch = userUrl.match(/systemusers\(([^)]+)\)/);
                if (userIdMatch) {
                    const userId = userIdMatch[1];
                    await assignSystemAdminRole(accessToken, userId);
                }
            }
            
            return { success: true, message: 'Application user created with System Administrator role' };
        } else {
            const errorText = await createResp.text();
            console.error('Could not create application user:', createResp.status, errorText);
            return { success: false, message: `Could not create application user: ${errorText}` };
        }
        
    } catch (error) {
        console.error('Error creating application user:', error);
        return { success: false, message: error.message };
    }
}

async function assignSystemAdminRole(accessToken, userId) {
    try {
        // Get the System Administrator role ID
        const roleResp = await fetch(
            `${currentOrgUrl}/api/data/v9.2/roles?$filter=name eq 'System Administrator'&$select=roleid`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json'
                }
            }
        );
        
        if (!roleResp.ok) {
            console.warn('Could not get System Administrator role');
            return;
        }
        
        const roleData = await roleResp.json();
        if (!roleData.value || roleData.value.length === 0) {
            console.warn('System Administrator role not found');
            return;
        }
        
        const roleId = roleData.value[0].roleid;
        
        // Associate the role with the user
        const associateResp = await fetch(
            `${currentOrgUrl}/api/data/v9.2/systemusers(${userId})/systemuserroles_association/$ref`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    '@odata.id': `${currentOrgUrl}/api/data/v9.2/roles(${roleId})`
                })
            }
        );
        
        if (associateResp.ok || associateResp.status === 204) {
            console.log('✅ System Administrator role assigned');
        } else {
            console.warn('Could not assign role:', associateResp.status);
        }
    } catch (error) {
        console.warn('Error assigning role:', error.message);
    }
}

async function disableSchedule() {
    const cfg = getSupabaseConfig();
    if (!cfg) return;
    
    const userEmail = getCurrentUserEmail();
    const envId = environmentId || '';
    
    if (!userEmail || !envId) return;
    
    try {
        await fetch(
            `${cfg.url}/rest/v1/update_schedules?user_email=eq.${encodeURIComponent(userEmail)}&environment_id=eq.${encodeURIComponent(envId)}`,
            {
                method: 'PATCH',
                headers: {
                    'apikey': cfg.key,
                    'Authorization': `Bearer ${cfg.key}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ enabled: false, updated_at: new Date().toISOString() })
            }
        );
        updateScheduleStatus(null);
    } catch (error) {
        console.warn('Failed to disable schedule:', error.message);
    }
}

// UI Helpers
function showLoading(message, details) {
    const overlay = document.getElementById('loadingOverlay');
    document.getElementById('loadingMessage').textContent = message;
    document.getElementById('loadingDetails').textContent = details || '';
    overlay.classList.remove('hidden');
    overlay.style.display = 'flex';
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    overlay.classList.add('hidden');
    overlay.style.display = 'none';
}

function showError(message) {
    hideLoading();
    // Use 'body' for HTML content, 'message' for plain text
    if (message.includes('<') && message.includes('>')) {
        showModal({ title: 'Error', body: message, type: 'danger', confirmOnly: true });
    } else {
        showModal({ title: 'Error', message: message, type: 'danger', confirmOnly: true });
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
}

// ── Custom Modal System ──────────────────────────────────────────────
let _modalResolve = null;

/**
 * Show a custom modal dialog. Returns a Promise<boolean>.
 * Options:
 *   title     – Modal title
 *   message   – Text or HTML message
 *   body      – Full HTML body (overrides message)
 *   type      – 'info' | 'warning' | 'success' | 'danger' | 'update'
 *   icon      – FontAwesome icon class (auto-selected from type if omitted)
 *   okText    – OK button text (default "OK")
 *   cancelText– Cancel button text (default "Cancel")
 *   okClass   – Extra class for OK button (e.g. 'btn-success-modal')
 *   confirmOnly – If true, hide Cancel button (alert-style)
 */
function showModal(opts) {
    return new Promise(resolve => {
        _modalResolve = resolve;
        
        const overlay = document.getElementById('customModal');
        const iconWrap = document.getElementById('modalIconWrap');
        const icon = document.getElementById('modalIcon');
        const title = document.getElementById('modalTitle');
        const body = document.getElementById('modalBody');
        const okBtn = document.getElementById('modalOkBtn');
        const cancelBtn = document.getElementById('modalCancelBtn');
        
        const typeIcons = {
            info: 'fas fa-info-circle',
            warning: 'fas fa-exclamation-triangle',
            success: 'fas fa-check-circle',
            danger: 'fas fa-times-circle',
            update: 'fas fa-arrow-circle-up'
        };
        
        const t = opts.type || 'info';
        iconWrap.className = 'modal-icon-wrap icon-' + t;
        icon.className = opts.icon || typeIcons[t] || typeIcons.info;
        title.textContent = opts.title || 'Notice';
        
        if (opts.body) {
            body.innerHTML = opts.body;
        } else {
            body.innerHTML = '<p class="mb-0">' + escapeHtml(opts.message || '') + '</p>';
        }
        
        okBtn.textContent = opts.okText || 'OK';
        okBtn.className = 'btn btn-modal-ok' + (opts.okClass ? ' ' + opts.okClass : '');
        cancelBtn.textContent = opts.cancelText || 'Cancel';
        cancelBtn.style.display = opts.confirmOnly ? 'none' : '';
        
        overlay.style.display = 'flex';
    });
}

function closeModal(result) {
    document.getElementById('customModal').style.display = 'none';
    if (_modalResolve) {
        _modalResolve(result);
        _modalResolve = null;
    }
}

/**
 * Helper: show a confirm modal for updating apps.
 * @param {Array} appsToUpdate – array of {name, version, latestVersion}
 * @returns Promise<boolean>
 */
function showUpdateConfirm(appsToUpdate) {
    let listHtml = '<ul class="update-list">';
    for (const app of appsToUpdate) {
        listHtml += '<li>';
        listHtml += '<span class="app-label" title="' + escapeHtml(app.name) + '">' + escapeHtml(app.name) + '</span>';
        listHtml += '<span class="version-badge">' + escapeHtml(app.version) + '<span class="arrow">→</span>' + escapeHtml(app.latestVersion || 'latest') + '</span>';
        listHtml += '</li>';
    }
    listHtml += '</ul>';
    
    const bodyHtml = '<p class="modal-message">The following ' + appsToUpdate.length + ' app' + (appsToUpdate.length !== 1 ? 's' : '') + ' will be updated:</p>' + listHtml;
    
    return showModal({
        title: 'Update Apps',
        body: bodyHtml,
        type: 'update',
        okText: 'Update All',
        okClass: 'btn-success-modal',
        cancelText: 'Cancel'
    });
}

/**
 * Helper: show a simple alert modal (no Cancel button).
 */
function showAlert(title, message, type) {
    return showModal({ title, message, type: type || 'info', confirmOnly: true });
}

window.updateSingleApp = updateSingleApp;
window.installApp = installApp;
window.closeModal = closeModal;
