// Global variables
let msalInstance = null;
let ppToken = null; // Power Platform API token
let environmentId = null;
let currentOrgUrl = null;
let apps = [];
let allEnvironments = []; // Cached list of all environments
let selectedApps = new Set(); // Multi-select tracking

// Supabase config for usage tracking
const SUPABASE_URL = 'https://fpekzltxukikaixebeeu.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZwZWt6bHR4dWtpa2FpeGViZWV1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzA0MDU0ODEsImV4cCI6MjA4NTk4MTQ4MX0.uH4JgKbf_-Al_iArzEy6UZ3edJNzFSCBVlMNI04li0Y';

// MSAL Configuration
function createMsalConfig(tenantId, clientId) {
    return {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
            redirectUri: window.location.origin + window.location.pathname.replace(/\/[^\/]*\.[^\/]*$/, '/'),
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false,
        },
    };
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM Content Loaded');
    hideLoading();
    
    if (typeof msal === 'undefined') {
        alert('Error: MSAL library failed to load.');
        return;
    }
    
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
    
    // Try to auto-login if we have saved credentials and a cached MSAL session
    tryAutoLogin();
    
    console.log('App initialized');
});

// Load saved credentials
function loadSavedCredentials() {
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            const orgUrlEl = document.getElementById('orgUrl');
            if (orgUrlEl) orgUrlEl.value = creds.orgUrl || creds.organizationId || creds.environmentId || '';
            document.getElementById('tenantId').value = creds.tenantId || '';
            document.getElementById('clientId').value = creds.clientId || '';
            document.getElementById('rememberMe').checked = true;
        } catch (e) {}
    }
}

// Auto-login: silently resume session if MSAL tokens are cached
async function tryAutoLogin() {
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (!savedCreds) return;

    let creds;
    try {
        creds = JSON.parse(savedCreds);
    } catch (e) {
        return;
    }

    const orgUrlValue = creds.orgUrl || creds.organizationId || creds.environmentId || '';
    const tenantId = creds.tenantId || '';
    const clientId = creds.clientId || '';
    if (!tenantId || !clientId || !orgUrlValue) return;

    try {
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        await msalInstance.handleRedirectPromise(); // Clear any stale redirect state

        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            // No cached session, user must log in manually
            msalInstance = null;
            return;
        }

        showLoading('Reconnecting...', 'Restoring your session');

        // Try to silently acquire Power Platform token
        const ppRequest = { scopes: ['https://api.powerplatform.com/.default'], account: accounts[0] };
        const ppResult = await msalInstance.acquireTokenSilent(ppRequest);
        ppToken = ppResult.accessToken;

        // Try to silently acquire BAP token (no interactive fallback during auto-login)
        const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account: accounts[0] };
        await msalInstance.acquireTokenSilent(bapRequest);

        showLoading('Reconnecting...', 'Resolving environment');

        // Normalize org URL
        let normalizedOrgUrl = orgUrlValue;
        if (!normalizedOrgUrl.startsWith('https://')) {
            normalizedOrgUrl = 'https://' + normalizedOrgUrl;
        }
        normalizedOrgUrl = normalizedOrgUrl.replace(/\/+$/, '');

        environmentId = await resolveOrgUrlToEnvironmentId(normalizedOrgUrl);
        if (!environmentId) {
            throw new Error('Could not resolve environment');
        }

        currentOrgUrl = normalizedOrgUrl;

        showLoading('Reconnecting...', 'Loading environment details');
        await getEnvironmentName();

        hideLoading();

        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appsSection').classList.remove('hidden');

        await loadApplications();

        console.log('Auto-login successful for', accounts[0].username);
    } catch (e) {
        // Silent acquisition failed — token expired or revoked, user must log in again
        console.log('Auto-login failed, user will log in manually:', e.message);
        hideLoading();
        msalInstance = null;
        ppToken = null;
        environmentId = null;
    }
}

// Handle authentication
async function handleAuthentication(event) {
    event.preventDefault();
    
    let orgUrlValue = document.getElementById('orgUrl').value.trim();
    const tenantId = document.getElementById('tenantId').value.trim();
    const clientId = document.getElementById('clientId').value.trim();
    const rememberMe = document.getElementById('rememberMe').checked;
    
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(tenantId) || !guidRegex.test(clientId)) {
        showError('Invalid GUID format. Tenant ID and Client ID must be valid GUIDs.');
        return;
    }
    
    // Normalize the org URL
    if (!orgUrlValue.startsWith('https://')) {
        orgUrlValue = 'https://' + orgUrlValue;
    }
    orgUrlValue = orgUrlValue.replace(/\/+$/, ''); // remove trailing slashes
    
    if (!orgUrlValue.includes('.dynamics.com')) {
        showError('Invalid Organization URL. It should look like https://yourorg.crm.dynamics.com');
        return;
    }
    
    if (rememberMe) {
        localStorage.setItem('d365_app_updater_creds', JSON.stringify({ orgUrl: orgUrlValue, tenantId, clientId }));
    }
    
    try {
        showLoading('Authenticating...', 'Connecting to Microsoft');
        
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        await msalInstance.handleRedirectPromise(); // Clear any stale redirect state
        
        // Sign in and consent to both APIs
        showLoading('Authenticating...', 'Signing in to Microsoft');
        const accounts = msalInstance.getAllAccounts();
        let account;
        
        if (accounts.length > 0) {
            account = accounts[0];
        } else {
            const loginResult = await msalInstance.loginPopup({
                scopes: ['https://api.powerplatform.com/.default'],
                extraScopesToConsent: ['https://api.bap.microsoft.com/.default']
            });
            account = loginResult.account;
        }
        
        // Get Power Platform API token (silent - consent already granted)
        showLoading('Authenticating...', 'Getting Power Platform API access');
        const ppRequest = { scopes: ['https://api.powerplatform.com/.default'], account };
        const ppResult = await msalInstance.acquireTokenSilent(ppRequest);
        ppToken = ppResult.accessToken;
        
        console.log('Power Platform API token acquired');
        
        // Eagerly acquire BAP token (popup fallback if consent wasn't granted)
        showLoading('Authenticating...', 'Getting BAP API access');
        const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account };
        try {
            await msalInstance.acquireTokenSilent(bapRequest);
        } catch (e) {
            console.log('BAP silent failed, requesting consent via popup...');
            await msalInstance.acquireTokenPopup(bapRequest);
        }
        
        console.log('BAP API token acquired');
        
        // Resolve Organization URL to Environment ID via BAP API
        showLoading('Authenticating...', 'Resolving Organization URL to Environment...');
        environmentId = await resolveOrgUrlToEnvironmentId(orgUrlValue);
        
        if (!environmentId) {
            throw new Error('Could not find a Power Platform environment matching URL: ' + orgUrlValue + '. Make sure you have admin access to the environment.');
        }
        
        console.log('Resolved Org URL', orgUrlValue, '→ Environment ID:', environmentId);
        currentOrgUrl = orgUrlValue;
        
        // Get environment name from BAP API (optional, just for display)
        showLoading('Loading...', 'Getting environment details');
        await getEnvironmentName();
        
        hideLoading();
        
        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appsSection').classList.remove('hidden');
        
        await loadApplications();
        
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
        // Fallback: try popup for consent (may be blocked if not user-initiated)
        try {
            const result = await msalInstance.acquireTokenPopup(bapRequest);
            return result.accessToken;
        } catch (popupError) {
            throw new Error('BAP API access denied. Please re-run the Setup Wizard to grant admin consent, or allow popups for this site.');
        }
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
                msalInstance.logoutPopup({ account: accounts[0] }).catch(() => {});
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
    showModal({ title: 'Error', message: message, type: 'danger', confirmOnly: true });
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
