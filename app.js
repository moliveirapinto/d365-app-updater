// Global variables
let msalInstance = null;
let accessToken = null;
let ppToken = null; // Power Platform API token
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
    document.getElementById('reinstallAllBtn').addEventListener('click', reinstallAllApps);
    
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
        
        // Get Power Platform API token
        showLoading('Authenticating...', 'Getting Power Platform API access');
        const ppRequest = { scopes: ['https://api.powerplatform.com/.default'], account: msalInstance.getAllAccounts()[0] };
        try {
            const ppResult = await msalInstance.acquireTokenSilent(ppRequest);
            ppToken = ppResult.accessToken;
        } catch (e) {
            const ppResult = await msalInstance.acquireTokenPopup(ppRequest);
            ppToken = ppResult.accessToken;
        }
        
        console.log('Power Platform API token acquired');
        
        // Get environment ID from BAP API
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
    // First get BAP token to list environments
    const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account: msalInstance.getAllAccounts()[0] };
    let bapToken;
    try {
        const bapResult = await msalInstance.acquireTokenSilent(bapRequest);
        bapToken = bapResult.accessToken;
    } catch (e) {
        const bapResult = await msalInstance.acquireTokenPopup(bapRequest);
        bapToken = bapResult.accessToken;
    }
    
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

// Compare two version strings (e.g., "1.2.3.4" vs "1.2.3.5")
function compareVersions(v1, v2) {
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

// Load applications from Power Platform API
async function loadApplications() {
    showLoading('Loading applications...', 'Fetching from Power Platform');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>';
    
    try {
        // Use the correct Power Platform API endpoint
        const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages?api-version=2022-03-01-preview`;
        
        console.log('Fetching apps from:', url);
        
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${ppToken}` }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Error:', response.status, errorText);
            throw new Error('Failed to fetch apps: ' + response.status);
        }
        
        const data = await response.json();
        console.log('Apps response:', data.value?.length, 'apps');
        
        // Separate installed apps from catalog apps
        const allApps = data.value || [];
        const installedApps = allApps.filter(a => a.state === 'Installed' || a.instancePackageId);
        const catalogApps = allApps.filter(a => a.state === 'None' || a.state === 'InstallAvailable');
        
        console.log('Installed apps:', installedApps.length);
        console.log('Catalog apps:', catalogApps.length);
        
        // Build a map of catalog apps by applicationId (keep the highest version)
        const catalogMap = new Map();
        for (const app of catalogApps) {
            if (app.applicationId) {
                const existing = catalogMap.get(app.applicationId);
                if (!existing || compareVersions(app.version, existing.version) > 0) {
                    catalogMap.set(app.applicationId, app);
                }
            }
        }
        
        // Process installed apps and check for updates
        apps = installedApps.map(app => {
            const catalogVersion = catalogMap.get(app.applicationId);
            let hasUpdate = false;
            let latestVersion = null;
            let catalogUniqueName = null;
            
            // Check if there's a newer version in the catalog
            if (catalogVersion && compareVersions(catalogVersion.version, app.version) > 0) {
                hasUpdate = true;
                latestVersion = catalogVersion.version;
                catalogUniqueName = catalogVersion.uniqueName;
            }
            
            return {
                id: app.id,
                uniqueName: app.uniqueName,
                catalogUniqueName: catalogUniqueName, // Use catalog uniqueName for install
                name: app.localizedName || app.applicationName || app.uniqueName || 'Unknown',
                version: app.version || 'Unknown',
                latestVersion: latestVersion,
                state: app.state || 'Installed',
                hasUpdate: hasUpdate,
                publisher: app.publisherName || 'Microsoft',
                description: app.applicationDescription || '',
                learnMoreUrl: app.learnMoreUrl || null,
                instancePackageId: app.instancePackageId,
                applicationId: app.applicationId
            };
        });
        
        // Also add apps from catalog that are not installed (for browse/install)
        const installedAppIds = new Set(installedApps.map(a => a.applicationId).filter(Boolean));
        const notInstalledApps = [];
        for (const [appId, app] of catalogMap) {
            if (!installedAppIds.has(appId)) {
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
        
        console.log('Apps with updates:', apps.filter(a => a.hasUpdate).length);
        
        // Sort: updates first, then installed, then available - all alphabetically
        apps.sort((a, b) => {
            if (a.hasUpdate && !b.hasUpdate) return -1;
            if (!a.hasUpdate && b.hasUpdate) return 1;
            return a.name.localeCompare(b.name);
        });
        
        // Store not-installed apps for browsing (sorted alphabetically)
        notInstalledApps.sort((a, b) => a.name.localeCompare(b.name));
        window.availableApps = notInstalledApps;
        
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
    
    // Show only apps with updates or installed apps (not the full catalog)
    const installedOrUpdatable = apps.filter(a => a.hasUpdate || a.instancePackageId);
    const appsToShow = installedOrUpdatable.length > 0 ? installedOrUpdatable : apps.slice(0, 50);
    
    let html = '';
    
    // Add info message about installed apps
    if (appsWithUpdates.length === 0 && installedOrUpdatable.length > 0) {
        html += '<div class="alert alert-info mb-3">';
        html += '<i class="fas fa-info-circle me-2"></i>';
        html += '<strong>Note:</strong> Showing ' + installedOrUpdatable.length + ' installed apps. ';
        html += 'Click "Reinstall" to trigger an update check for any app.';
        html += '</div>';
    }
    
    for (const app of appsToShow) {
        const stateClass = app.hasUpdate ? 'success' : 'secondary';
        const stateIcon = app.hasUpdate ? 'arrow-circle-up' : 'check-circle';
        const stateText = app.hasUpdate ? 'Update Available' : (app.instancePackageId ? 'Installed' : 'Available');
        
        html += '<div class="app-card">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-6">';
        html += '<div class="app-name"><i class="fas fa-cube me-2"></i>' + escapeHtml(app.name) + '</div>';
        html += '<div class="app-version mt-2">';
        html += '<i class="fas fa-tag"></i> Version: <strong>' + escapeHtml(app.version) + '</strong>';
        if (app.hasUpdate && app.latestVersion) {
            html += ' → <strong class="text-success">' + escapeHtml(app.latestVersion) + '</strong>';
        }
        html += '</div>';
        html += '<div class="text-muted small mt-1"><i class="fas fa-building"></i> ' + escapeHtml(app.publisher) + '</div>';
        html += '</div>';
        html += '<div class="col-md-3 text-center">';
        html += '<span class="badge bg-' + stateClass + '"><i class="fas fa-' + stateIcon + '"></i> ' + stateText + '</span>';
        html += '</div>';
        html += '<div class="col-md-3 text-end">';
        if (app.hasUpdate) {
            html += '<button class="btn btn-success btn-sm" onclick="updateSingleApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-download"></i> Update</button>';
        } else if (!app.instancePackageId) {
            html += '<button class="btn btn-primary btn-sm" onclick="installApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-plus"></i> Install</button>';
        } else {
            // Show reinstall button for installed apps
            html += '<button class="btn btn-outline-primary btn-sm" onclick="reinstallApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-sync"></i> Reinstall</button>';
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
    
    if (!confirm('Install update for "' + app.name + '"?\n\nCurrent: ' + app.version + '\nNew: ' + app.latestVersion)) {
        return;
    }
    
    showLoading('Installing update...', app.name);
    
    try {
        // Use the catalog's uniqueName if available (for updates), otherwise use the installed app's uniqueName
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
            throw new Error('Install failed: ' + response.status + ' - ' + errorText);
        }
        
        hideLoading();
        alert('Update started for ' + app.name + '!\n\nThe update is running in the background and may take several minutes.');
        
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Update error:', error);
        showError('Failed to update: ' + error.message);
    }
}

// Install an app
async function installApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    if (!confirm('Install "' + app.name + '"?')) {
        return;
    }
    
    showLoading('Installing...', app.name);
    
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
            throw new Error('Install failed: ' + response.status + ' - ' + errorText);
        }
        
        hideLoading();
        alert('Installation started for ' + app.name + '!\n\nThis may take several minutes.');
        
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Install error:', error);
        showError('Failed to install: ' + error.message);
    }
}

// Reinstall an already installed app (to apply any available updates)
async function reinstallApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    if (!confirm('Reinstall "' + app.name + '"?\n\nThis will check for and apply any available updates.\n\nCurrent version: ' + app.version)) {
        return;
    }
    
    showLoading('Reinstalling...', app.name);
    
    try {
        const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2022-03-01-preview`;
        
        console.log('Reinstalling:', app.name, 'using package:', app.uniqueName);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${ppToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error('Reinstall failed: ' + response.status + ' - ' + errorText);
        }
        
        hideLoading();
        alert('Reinstall/update started for ' + app.name + '!\n\nThe system will apply any available updates. This may take several minutes.');
        
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Reinstall error:', error);
        showError('Failed to reinstall: ' + error.message);
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
            // Use the catalog's uniqueName if available (for updates), otherwise use the installed app's uniqueName
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

// Reinstall all installed apps
async function reinstallAllApps() {
    const installedApps = apps.filter(a => a.instancePackageId);
    
    if (installedApps.length === 0) {
        alert('No installed apps found.');
        return;
    }
    
    if (!confirm('Reinstall all ' + installedApps.length + ' installed applications?\n\nThis will trigger update checks for each installed app.\nApps that are already current will remain unchanged.\n\nThis operation may take several minutes.')) {
        return;
    }
    
    showLoading('Reinstalling apps...', '0 of ' + installedApps.length);
    
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < installedApps.length; i++) {
        const app = installedApps[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + installedApps.length + ': ' + app.name;
        
        try {
            const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${app.uniqueName}/install?api-version=2022-03-01-preview`;
            
            console.log('Reinstalling:', app.name, 'using package:', app.uniqueName);
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${ppToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.ok) {
                successCount++;
            } else {
                failCount++;
                console.error('Failed to reinstall ' + app.name + ':', response.status);
            }
        } catch (error) {
            failCount++;
            console.error('Error reinstalling ' + app.name + ':', error);
        }
        
        // Small delay between requests to avoid rate limiting
        await new Promise(r => setTimeout(r, 1500));
    }
    
    hideLoading();
    
    if (failCount === 0) {
        alert('All ' + successCount + ' reinstall requests submitted successfully!\n\nUpdates are running in the background and may take several minutes.\n\nCheck the Power Platform Admin Center for progress.');
    } else {
        alert('Submitted ' + successCount + ' reinstall requests.\n' + failCount + ' failed.\n\nCheck the Power Platform Admin Center for details.');
    }
    
    await loadApplications();
}

// Logout
function handleLogout() {
    if (confirm('Logout?')) {
        accessToken = null;
        ppToken = null;
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
window.installApp = installApp;
