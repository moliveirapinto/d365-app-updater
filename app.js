// Global variables
let msalInstance = null;
let ppToken = null; // Power Platform API token
let environmentId = null;
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
            document.getElementById('environmentId').value = creds.environmentId || '';
            document.getElementById('tenantId').value = creds.tenantId || '';
            document.getElementById('clientId').value = creds.clientId || '';
            document.getElementById('rememberMe').checked = true;
        } catch (e) {}
    }
}

// Handle authentication
async function handleAuthentication(event) {
    event.preventDefault();
    
    const envIdValue = document.getElementById('environmentId').value.trim();
    const tenantId = document.getElementById('tenantId').value.trim();
    const clientId = document.getElementById('clientId').value.trim();
    const rememberMe = document.getElementById('rememberMe').checked;
    
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(tenantId) || !guidRegex.test(clientId) || !guidRegex.test(envIdValue)) {
        showError('Invalid GUID format. All IDs must be valid GUIDs.');
        return;
    }
    
    // Use the Environment ID directly
    environmentId = envIdValue;
    
    if (rememberMe) {
        localStorage.setItem('d365_app_updater_creds', JSON.stringify({ environmentId: envIdValue, tenantId, clientId }));
    }
    
    try {
        showLoading('Authenticating...', 'Connecting to Microsoft');
        
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        
        // Get Power Platform API token
        showLoading('Authenticating...', 'Getting Power Platform API access');
        const accounts = msalInstance.getAllAccounts();
        const ppRequest = { scopes: ['https://api.powerplatform.com/.default'], account: accounts[0] };
        
        let ppResult;
        if (accounts.length > 0) {
            try {
                ppResult = await msalInstance.acquireTokenSilent(ppRequest);
            } catch (e) {
                ppResult = await msalInstance.acquireTokenPopup(ppRequest);
            }
        } else {
            ppResult = await msalInstance.acquireTokenPopup(ppRequest);
        }
        ppToken = ppResult.accessToken;
        
        console.log('Power Platform API token acquired');
        console.log('Using Environment ID:', environmentId);
        
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

// Get environment name from BAP API (for display purposes)
async function getEnvironmentName() {
    // First get BAP token
    const bapRequest = { scopes: ['https://api.bap.microsoft.com/.default'], account: msalInstance.getAllAccounts()[0] };
    let bapToken;
    try {
        const bapResult = await msalInstance.acquireTokenSilent(bapRequest);
        bapToken = bapResult.accessToken;
    } catch (e) {
        const bapResult = await msalInstance.acquireTokenPopup(bapRequest);
        bapToken = bapResult.accessToken;
    }
    
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
        
        // Add timeout to the fetch
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 60000); // 60 second timeout
        
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${ppToken}` },
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
        console.log('Response status:', response.status);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Error:', response.status, errorText);
            throw new Error('Failed to fetch apps: ' + response.status);
        }
        
        const data = await response.json();
        const allApps = data.value || [];
        console.log('Apps response:', allApps.length, 'apps');
        
        // Debug: Log all unique states and check for update-related fields
        const states = [...new Set(allApps.map(a => a.state))];
        console.log('All states found:', states);
        
        // Check first few apps for any update-related fields
        if (allApps.length > 0) {
            console.log('Sample app fields:', Object.keys(allApps[0]));
            // Look for installed apps that might have update info
            const sampleInstalled = allApps.find(a => a.state === 'Installed');
            if (sampleInstalled) {
                console.log('Sample installed app:', JSON.stringify(sampleInstalled, null, 2));
            }
        }
        
        // Separate installed apps from catalog apps (include all non-installed states as catalog)
        const installedApps = allApps.filter(a => a.state === 'Installed' || a.instancePackageId);
        const catalogApps = allApps.filter(a => a.state !== 'Installed' && !a.instancePackageId);
        
        console.log('Installed apps:', installedApps.length);
        console.log('Catalog apps (available):', catalogApps.length);
        
        // Build a map of ALL catalog apps by applicationId (keep the highest version)
        const catalogMap = new Map();
        for (const app of catalogApps) {
            if (app.applicationId) {
                const existing = catalogMap.get(app.applicationId);
                if (!existing || compareVersions(app.version, existing.version) > 0) {
                    catalogMap.set(app.applicationId, app);
                }
            }
        }
        
        // Also build a map by uniqueName for apps that share the same base name
        const catalogByName = new Map();
        for (const app of catalogApps) {
            if (app.uniqueName) {
                const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                const existing = catalogByName.get(baseName);
                if (!existing || compareVersions(app.version, existing.version) > 0) {
                    catalogByName.set(baseName, app);
                }
            }
        }
        
        // Process installed apps and check for updates
        let updatesFound = 0;
        apps = installedApps.map(app => {
            let hasUpdate = false;
            let latestVersion = null;
            let catalogUniqueName = null;
            
            // Method 1: Check by applicationId
            const catalogVersion = catalogMap.get(app.applicationId);
            if (catalogVersion && compareVersions(catalogVersion.version, app.version) > 0) {
                hasUpdate = true;
                latestVersion = catalogVersion.version;
                catalogUniqueName = catalogVersion.uniqueName;
            }
            
            // Method 2: Check by uniqueName base (in case applicationId doesn't match)
            if (!hasUpdate && app.uniqueName) {
                const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                const byName = catalogByName.get(baseName);
                if (byName && compareVersions(byName.version, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = byName.version;
                    catalogUniqueName = byName.uniqueName;
                }
            }
            
            // Method 3: Check if the API directly provides update info
            if (!hasUpdate && app.updateAvailable) {
                hasUpdate = true;
            }
            
            if (hasUpdate) {
                updatesFound++;
                console.log('Update found for:', app.localizedName || app.uniqueName, app.version, '->', latestVersion);
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
                applicationId: app.applicationId
            };
        });
        
        console.log('Total updates found via version comparison:', updatesFound);
        
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
        
        console.log('Displaying applications...');
        displayApplications();
        hideLoading();
        console.log('Loading complete');
        
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
    
    const appsWithUpdates = apps.filter(a => a.hasUpdate);
    const installedApps = apps.filter(a => a.instancePackageId);
    document.getElementById('appCountText').textContent = installedApps.length + ' apps installed';
    document.getElementById('updateAllBtn').disabled = appsWithUpdates.length === 0;
    
    // Show installed apps
    const installedOrUpdatable = apps.filter(a => a.hasUpdate || a.instancePackageId);
    const appsToShow = installedOrUpdatable.length > 0 ? installedOrUpdatable : apps.slice(0, 50);
    
    let html = '';
    
    // Add info message
    if (installedOrUpdatable.length > 0) {
        html += '<div class="alert alert-success mb-3">';
        html += '<i class="fas fa-check-circle me-2"></i>';
        html += 'Click <strong>"Update All Apps"</strong> to apply any available updates, or update apps individually.';
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
            // Show Update button only when we detect an update is available
            html += '<button class="btn btn-success btn-sm" onclick="updateSingleApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-download"></i> Update</button>';
        } else if (!app.instancePackageId) {
            // Show Install button for apps not installed
            html += '<button class="btn btn-primary btn-sm" onclick="installApp(\'' + escapeHtml(app.uniqueName) + '\')"><i class="fas fa-plus"></i> Install</button>';
        } else {
            // Installed and up to date - show checkmark
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

// Update an installed app
async function reinstallApp(uniqueName) {
    const app = apps.find(a => a.uniqueName === uniqueName);
    if (!app) return;
    
    if (!confirm('Update "' + app.name + '"?\n\nCurrent version: ' + app.version)) {
        return;
    }
    
    showLoading('Updating...', app.name);
    
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
            throw new Error('Update failed: ' + response.status + ' - ' + errorText);
        }
        
        hideLoading();
        alert('Update started for ' + app.name + '!\n\nThis may take several minutes.');
        
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

// Update all installed apps
async function reinstallAllApps() {
    const installedApps = apps.filter(a => a.instancePackageId);
    
    if (installedApps.length === 0) {
        alert('No installed apps found.');
        return;
    }
    
    if (!confirm('Update all ' + installedApps.length + ' installed apps?\n\nThis will apply available updates to all apps.')) {
        return;
    }
    
    showLoading('Updating apps...', '0 of ' + installedApps.length);
    
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < installedApps.length; i++) {
        const app = installedApps[i];
        document.getElementById('loadingDetails').textContent = (i + 1) + ' of ' + installedApps.length + ': ' + app.name;
        
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
        
        // Small delay between requests to avoid rate limiting
        await new Promise(r => setTimeout(r, 1500));
    }
    
    hideLoading();
    
    if (failCount === 0) {
        alert('All ' + successCount + ' update requests submitted!\n\nUpdates are running in the background and may take several minutes.');
    } else {
        alert('Submitted ' + successCount + ' updates.\n' + failCount + ' failed.\n\nCheck the Power Platform Admin Center for details.');
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
