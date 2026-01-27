// Global variables
let msalInstance = null;
let accessToken = null;
let bapToken = null;
let environmentId = null;
let orgUrl = null;
let apps = [];

// MSAL Configuration
function createMsalConfig(tenantId, clientId) {
    return {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
            redirectUri: window.location.origin,
        },
        cache: {
            cacheLocation: 'sessionStorage',
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
    
    console.log('App initialized');
});

// Load saved credentials
function loadSavedCredentials() {
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            document.getElementById('orgUrl').value = creds.orgUrl || '';
            document.getElementById('tenantId').value = creds.tenantId || '';
            document.getElementById('clientId').value = creds.clientId || '';
            document.getElementById('rememberMe').checked = true;
        } catch (e) {}
    }
}

// Handle authentication
async function handleAuthentication(event) {
    event.preventDefault();
    
    const orgUrlValue = document.getElementById('orgUrl').value.trim();
    const tenantId = document.getElementById('tenantId').value.trim();
    const clientId = document.getElementById('clientId').value.trim();
    const rememberMe = document.getElementById('rememberMe').checked;
    
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(tenantId) || !guidRegex.test(clientId)) {
        showError('Invalid GUID format');
        return;
    }
    
    orgUrl = orgUrlValue.replace(/\/$/, '');
    
    if (rememberMe) {
        localStorage.setItem('d365_app_updater_creds', JSON.stringify({ orgUrl, tenantId, clientId }));
    }
    
    try {
        showLoading('Authenticating...', 'Connecting to Microsoft');
        
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        
        // Get D365 token first
        const accounts = msalInstance.getAllAccounts();
        const d365Request = { scopes: [`${orgUrl}/.default`], account: accounts[0] };
        
        let authResult;
        if (accounts.length > 0) {
            try {
                authResult = await msalInstance.acquireTokenSilent(d365Request);
            } catch (e) {
                authResult = await msalInstance.acquireTokenPopup(d365Request);
            }
        } else {
            authResult = await msalInstance.acquireTokenPopup(d365Request);
        }
        accessToken = authResult.accessToken;
        
        // Get BAP token for Power Platform API
        showLoading('Authenticating...', 'Getting Power Platform API access');
        const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account: msalInstance.getAllAccounts()[0] };
        try {
            const bapResult = await msalInstance.acquireTokenSilent(bapRequest);
            bapToken = bapResult.accessToken;
        } catch (e) {
            const bapResult = await msalInstance.acquireTokenPopup(bapRequest);
            bapToken = bapResult.accessToken;
        }
        
        console.log('BAP token acquired');
        
        // Get environment ID
        showLoading('Loading...', 'Finding your environment');
        await findEnvironment();
        
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

// Find the Power Platform environment matching the org URL
async function findEnvironment() {
    const response = await fetch('https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2021-04-01', {
        headers: { 'Authorization': `Bearer ${bapToken}` }
    });
    
    if (!response.ok) {
        throw new Error('Failed to get environments: ' + response.status);
    }
    
    const data = await response.json();
    console.log('Environments:', data.value?.length);
    
    // Find environment matching our org URL
    const orgHost = new URL(orgUrl).hostname.toLowerCase();
    
    for (const env of data.value || []) {
        const instanceUrl = env.properties?.linkedEnvironmentMetadata?.instanceUrl;
        if (instanceUrl) {
            const envHost = new URL(instanceUrl).hostname.toLowerCase();
            if (envHost === orgHost) {
                environmentId = env.name;
                document.getElementById('environmentName').textContent = env.properties?.displayName || env.name;
                console.log('Found environment:', environmentId);
                return;
            }
        }
    }
    
    throw new Error('Could not find Power Platform environment for ' + orgUrl);
}

// Load applications from Power Platform API
async function loadApplications() {
    showLoading('Loading applications...', 'Fetching from Power Platform');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>';
    
    try {
        // Use the correct endpoint for Dynamics 365 apps
        const url = `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environmentId}/applicationPackages?api-version=2016-11-01`;
        
        console.log('Fetching apps from:', url);
        
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${bapToken}` }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Error:', response.status, errorText);
            throw new Error('Failed to fetch apps: ' + response.status);
        }
        
        const data = await response.json();
        console.log('Apps response:', data);
        
        apps = (data.value || []).map(app => {
            const props = app.properties || {};
            return {
                id: app.id || app.name,
                uniqueName: props.uniqueName || props.applicationId || app.name,
                name: props.localizedDisplayName || props.displayName || props.uniqueName || 'Unknown',
                version: props.version || props.applicationVersion || 'Unknown',
                latestVersion: props.latestVersion || null,
                state: props.state || 'Unknown',
                hasUpdate: props.state === 'UpdateAvailable' || (props.latestVersion && props.latestVersion !== props.version),
                publisher: props.publisherName || props.publisherId || 'Microsoft',
                learnMoreUrl: props.learnMoreUrl || null
            };
        });
        
        // Sort: updates first, then alphabetically
        apps.sort((a, b) => {
            if (a.hasUpdate && !b.hasUpdate) return -1;
            if (!a.hasUpdate && b.hasUpdate) return 1;
            return a.name.localeCompare(b.name);
        });
        
        displayApplications();
        hideLoading();
        
    } catch (error) {
        hideLoading();
        console.error('Error:', error);
        appsList.innerHTML = '<div class="alert alert-danger">Failed to load: ' + error.message + '</div>';
    }
}

// Display applications
function displayApplications() {
    const appsList = document.getElementById('appsList');
    
    if (apps.length === 0) {
        appsList.innerHTML = '<div class="text-center py-5"><p class="text-muted">No applications found.</p></div>';
        return;
    }
    
    const appsWithUpdates = apps.filter(a => a.hasUpdate);
    document.getElementById('updateCount').textContent = appsWithUpdates.length;
    document.getElementById('updateAllBtn').disabled = appsWithUpdates.length === 0;
    
    let html = '';
    
    for (const app of apps) {
        html += '<div class="app-card">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-6">';
        html += '<div class="app-name"><i class="fas fa-cube me-2"></i>' + escapeHtml(app.name) + '</div>';
        html += '<div class="app-version mt-2">';
        html += '<i class="fas fa-tag"></i> Installed: <strong>' + escapeHtml(app.version) + '</strong>';
        if (app.hasUpdate && app.latestVersion) {
            html += '<br><i class="fas fa-arrow-up text-success"></i> Available: <strong class="text-success">' + escapeHtml(app.latestVersion) + '</strong>';
        }
        html += '</div>';
        html += '<div class="text-muted small mt-1"><i class="fas fa-building"></i> ' + escapeHtml(app.publisher) + '</div>';
        html += '</div>';
        html += '<div class="col-md-3 text-center">';
        if (app.hasUpdate) {
            html += '<span class="badge-update"><i class="fas fa-arrow-circle-up"></i> Update Available</span>';
        } else {
            html += '<span class="badge-current"><i class="fas fa-check-circle"></i> Up to Date</span>';
        }
        html += '</div>';
        html += '<div class="col-md-3 text-end">';
        if (app.hasUpdate) {
            html += '<button class="btn btn-success btn-sm" onclick="updateSingleApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-download"></i> Update</button>';
        } else {
            html += '<button class="btn btn-outline-secondary btn-sm" disabled><i class="fas fa-check"></i> Current</button>';
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
    
    if (!confirm('Install update for "' + app.name + '"?\n\nThis will update from v' + app.version + ' to v' + app.latestVersion)) {
        return;
    }
    
    showLoading('Installing update...', app.name);
    
    try {
        const url = `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2016-11-01`;
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${bapToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error('Install failed: ' + response.status + ' - ' + errorText);
        }
        
        hideLoading();
        alert('Update started for ' + app.name + '!\n\nThe update is now running in the background. It may take several minutes to complete.');
        
        // Refresh the list
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Update error:', error);
        showError('Failed to update: ' + error.message);
    }
}

// Update all apps
async function updateAllApps() {
    const appsToUpdate = apps.filter(a => a.hasUpdate);
    
    if (appsToUpdate.length === 0) {
        alert('No updates available');
        return;
    }
    
    if (!confirm('Install updates for ' + appsToUpdate.length + ' applications?\n\nThis will update all apps with available updates.')) {
        return;
    }
    
    showLoading('Installing updates...', '0 of ' + appsToUpdate.length);
    
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + appsToUpdate.length + ': ' + app.name;
        
        try {
            const url = `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2016-11-01`;
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${bapToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.ok) {
                successCount++;
            } else {
                failCount++;
                console.error('Failed to update ' + app.name + ':', response.status);
            }
        } catch (error) {
            failCount++;
            console.error('Error updating ' + app.name + ':', error);
        }
        
        // Small delay between requests
        await new Promise(r => setTimeout(r, 1000));
    }
    
    hideLoading();
    
    if (failCount === 0) {
        alert('All ' + successCount + ' updates started successfully!\n\nUpdates are running in the background and may take several minutes.');
    } else {
        alert('Started ' + successCount + ' updates.\n' + failCount + ' failed.\n\nCheck the Power Platform Admin Center for details.');
    }
    
    await loadApplications();
}

// Logout
function handleLogout() {
    if (confirm('Logout?')) {
        accessToken = null;
        bapToken = null;
        environmentId = null;
        apps = [];
        
        if (msalInstance) {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.logoutPopup({ account: accounts[0] }).catch(() => {});
            }
        }
        
        document.getElementById('appsSection').classList.add('hidden');
        document.getElementById('authSection').classList.remove('hidden');
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
    alert('Error: ' + message);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
}

window.updateSingleApp = updateSingleApp;
