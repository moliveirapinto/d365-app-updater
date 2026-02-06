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
        
        // Resolve Organization URL to Environment ID via BAP API
        showLoading('Authenticating...', 'Resolving Organization URL to Environment...');
        environmentId = await resolveOrgUrlToEnvironmentId(orgUrlValue);
        
        if (!environmentId) {
            throw new Error('Could not find a Power Platform environment matching URL: ' + orgUrlValue + '. Make sure you have admin access to the environment.');
        }
        
        console.log('Resolved Org URL', orgUrlValue, '→ Environment ID:', environmentId);
        
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
    
    for (const env of environments) {
        const instanceUrl = env.properties?.linkedEnvironmentMetadata?.instanceUrl;
        const envName = env.properties?.displayName || env.name;
        
        if (instanceUrl) {
            const normalizedInstance = instanceUrl.toLowerCase().replace(/\/+$/, '');
            console.log(`  Environment: ${envName} (${env.name}), instanceUrl: ${instanceUrl}`);
            
            if (normalizedInstance === normalizedInput) {
                console.log('  ✓ Match found! Environment ID:', env.name);
                return env.name; // env.name is the Environment ID (GUID)
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
        const result = await msalInstance.acquireTokenPopup(bapRequest);
        return result.accessToken;
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
        
        // Debug: log unique states and sample data
        const states = [...new Set(allAppsRaw.map(a => a.state))];
        console.log('All states found in catalog:', states);
        if (installedAppsRaw.length > 0) {
            console.log('Sample installed app fields:', Object.keys(installedAppsRaw[0]));
            console.log('Sample installed app:', JSON.stringify(installedAppsRaw[0], null, 2));
        }
        
        // ── Step 3: Build version maps from ALL catalog entries ──────
        // Map by applicationId → keep highest version
        const catalogMapById = new Map();
        for (const app of allAppsRaw) {
            if (!app.applicationId) continue;
            const existing = catalogMapById.get(app.applicationId);
            if (!existing || compareVersions(app.version, existing.version) > 0) {
                catalogMapById.set(app.applicationId, app);
            }
        }
        
        // Map by uniqueName base → keep highest version (fallback matching)
        const catalogByName = new Map();
        for (const app of allAppsRaw) {
            if (!app.uniqueName) continue;
            const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
            const existing = catalogByName.get(baseName);
            if (!existing || compareVersions(app.version, existing.version) > 0) {
                catalogByName.set(baseName, app);
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
            
            // Check 1: Direct API fields that might indicate update availability
            if (app.updateAvailable || app.catalogVersion || app.availableVersion || app.latestVersion) {
                const directVersion = app.catalogVersion || app.availableVersion || app.latestVersion;
                if (directVersion && compareVersions(directVersion, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = directVersion;
                }
            }
            
            // Check 2: Compare with catalog entry by applicationId
            if (!hasUpdate && app.applicationId) {
                const catalogEntry = catalogMapById.get(app.applicationId);
                if (catalogEntry && compareVersions(catalogEntry.version, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = catalogEntry.version;
                    catalogUniqueName = catalogEntry.uniqueName;
                    console.log(`  [by appId] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                }
            }
            
            // Check 3: Compare with catalog entry by uniqueName base
            if (!hasUpdate && app.uniqueName) {
                const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                const byName = catalogByName.get(baseName);
                if (byName && compareVersions(byName.version, app.version) > 0) {
                    hasUpdate = true;
                    latestVersion = byName.version;
                    catalogUniqueName = byName.uniqueName;
                    console.log(`  [by name] ${app.localizedName || app.uniqueName}: ${app.version} → ${latestVersion}`);
                }
            }
            
            if (hasUpdate) updatesFound++;
            
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
        
        console.log('Updates found from PP API:', updatesFound);
        
        // ── Step 5: If PP API found zero updates, try BAP Admin API ──
        if (updatesFound === 0 && installedAppsRaw.length > 0) {
            console.log('No updates found via PP API. Trying BAP Admin API as fallback...');
            showLoading('Loading applications...', 'Checking for updates via Admin API...');
            
            try {
                const bapToken = await getBAPToken();
                const bapUrl = `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentId}/applicationPackages?api-version=2021-04-01`;
                const bapApps = await fetchAllPages(bapUrl, bapToken);
                
                console.log('BAP API returned:', bapApps.length, 'packages');
                if (bapApps.length > 0) {
                    console.log('BAP sample app fields:', Object.keys(bapApps[0]));
                    console.log('BAP sample app:', JSON.stringify(bapApps[0], null, 2));
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
                
                // Re-check installed apps against BAP data
                for (const app of apps) {
                    if (app.hasUpdate) continue;
                    
                    // Check direct fields from BAP response for this app
                    const bapInstalled = bapApps.find(b => 
                        (b.applicationId === app.applicationId) && 
                        (b.state === 'Installed' || b.instancePackageId)
                    );
                    if (bapInstalled) {
                        // Check if BAP provides update info directly
                        const directVer = bapInstalled.catalogVersion || bapInstalled.availableVersion || bapInstalled.latestVersion;
                        if (directVer && compareVersions(directVer, app.version) > 0) {
                            app.hasUpdate = true;
                            app.latestVersion = directVer;
                            updatesFound++;
                            console.log(`  [BAP direct] ${app.name}: ${app.version} → ${directVer}`);
                            continue;
                        }
                    }
                    
                    // Compare by applicationId
                    if (app.applicationId) {
                        const bapEntry = bapCatalogMap.get(app.applicationId);
                        if (bapEntry && compareVersions(bapEntry.version, app.version) > 0) {
                            app.hasUpdate = true;
                            app.latestVersion = bapEntry.version;
                            app.catalogUniqueName = bapEntry.uniqueName || app.uniqueName;
                            updatesFound++;
                            console.log(`  [BAP by appId] ${app.name}: ${app.version} → ${bapEntry.version}`);
                            continue;
                        }
                    }
                    
                    // Compare by uniqueName
                    if (app.uniqueName) {
                        const baseName = app.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
                        const bapByNameEntry = bapByName.get(baseName);
                        if (bapByNameEntry && compareVersions(bapByNameEntry.version, app.version) > 0) {
                            app.hasUpdate = true;
                            app.latestVersion = bapByNameEntry.version;
                            app.catalogUniqueName = bapByNameEntry.uniqueName || app.uniqueName;
                            updatesFound++;
                            console.log(`  [BAP by name] ${app.name}: ${app.version} → ${bapByNameEntry.version}`);
                        }
                    }
                }
                
                console.log('Updates found after BAP fallback:', updatesFound);
            } catch (bapError) {
                console.warn('BAP API fallback failed (non-critical):', bapError.message);
            }
        }
        
        // ── Step 6: Sort — updates first, then alphabetically ────────
        apps.sort((a, b) => {
            if (a.hasUpdate && !b.hasUpdate) return -1;
            if (!a.hasUpdate && b.hasUpdate) return 1;
            return a.name.localeCompare(b.name);
        });
        
        // Store not-installed apps for browsing
        const installedAppIds = new Set(installedAppsRaw.map(a => a.applicationId).filter(Boolean));
        const notInstalledApps = [];
        for (const [appId, app] of catalogMapById) {
            if (!installedAppIds.has(appId) && app.state !== 'Installed' && !app.instancePackageId) {
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
    
    const appsWithUpdates = apps.filter(a => a.hasUpdate);
    const installedApps = apps.filter(a => a.instancePackageId);
    const updateCount = appsWithUpdates.length;
    
    // Update summary text
    if (updateCount > 0) {
        document.getElementById('appCountText').innerHTML = 
            installedApps.length + ' apps installed &nbsp;|&nbsp; <span style="color: #28a745; font-weight: 600;">' + 
            updateCount + ' update' + (updateCount !== 1 ? 's' : '') + ' available</span>';
    } else {
        document.getElementById('appCountText').textContent = installedApps.length + ' apps installed — all up to date';
    }
    
    document.getElementById('updateAllBtn').disabled = updateCount === 0;
    
    // Show installed apps
    const installedOrUpdatable = apps.filter(a => a.hasUpdate || a.instancePackageId);
    const appsToShow = installedOrUpdatable.length > 0 ? installedOrUpdatable : apps.slice(0, 50);
    
    let html = '';
    
    // Update summary banner
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
    } else {
        html += '<div class="alert alert-success mb-3">';
        html += '<i class="fas fa-check-circle me-2"></i>';
        html += 'All installed applications are up to date.';
        html += '</div>';
    }
    
    for (const app of appsToShow) {
        const stateClass = app.hasUpdate ? 'success' : 'secondary';
        const stateIcon = app.hasUpdate ? 'arrow-circle-up' : 'check-circle';
        const stateText = app.hasUpdate ? 'Update Available' : (app.instancePackageId ? 'Installed' : 'Available');
        
        html += '<div class="app-card' + (app.hasUpdate ? ' border-success' : '') + '" style="' + (app.hasUpdate ? 'border-left: 4px solid #28a745; background: #f8fff8;' : '') + '">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-6">';
        html += '<div class="app-name"><i class="fas fa-cube me-2"></i>' + escapeHtml(app.name) + '</div>';
        html += '<div class="app-version mt-2">';
        html += '<i class="fas fa-tag"></i> Version: <strong>' + escapeHtml(app.version) + '</strong>';
        if (app.hasUpdate && app.latestVersion) {
            html += ' <i class="fas fa-long-arrow-alt-right text-success mx-1"></i> <strong class="text-success">' + escapeHtml(app.latestVersion) + '</strong>';
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
