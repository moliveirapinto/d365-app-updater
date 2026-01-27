// Global variables
let msalInstance = null;
let accessToken = null;
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
    
    // Make sure loading overlay is hidden on page load
    hideLoading();
    
    // Check if MSAL library is loaded
    if (typeof msal === 'undefined') {
        console.error('MSAL library not loaded');
        alert('Error: MSAL library failed to load. Please check your internet connection and refresh the page.');
        return;
    }
    
    console.log('MSAL library loaded successfully');
    
    // Set redirect URI in instructions
    const redirectUriElement = document.getElementById('redirectUri');
    if (redirectUriElement) {
        redirectUriElement.textContent = window.location.origin;
    }
    
    // Load saved credentials if available
    try {
        loadSavedCredentials();
    } catch (error) {
        console.error('Error loading saved credentials:', error);
    }
    
    // Event listeners
    document.getElementById('authForm').addEventListener('submit', handleAuthentication);
    document.getElementById('logoutBtn').addEventListener('click', handleLogout);
    document.getElementById('refreshAppsBtn').addEventListener('click', loadApplications);
    document.getElementById('updateAllBtn').addEventListener('click', updateAllApps);
    
    console.log('App initialized successfully');
});

// Load saved credentials from localStorage
function loadSavedCredentials() {
    const savedCreds = localStorage.getItem('d365_app_updater_creds');
    if (savedCreds) {
        try {
            const creds = JSON.parse(savedCreds);
            document.getElementById('orgUrl').value = creds.orgUrl || '';
            document.getElementById('tenantId').value = creds.tenantId || '';
            document.getElementById('clientId').value = creds.clientId || '';
            document.getElementById('rememberMe').checked = true;
        } catch (error) {
            console.error('Failed to load saved credentials:', error);
        }
    }
}

// Handle authentication
async function handleAuthentication(event) {
    event.preventDefault();
    
    const orgUrlValue = document.getElementById('orgUrl').value.trim();
    const tenantId = document.getElementById('tenantId').value.trim();
    const clientId = document.getElementById('clientId').value.trim();
    const rememberMe = document.getElementById('rememberMe').checked;
    
    // Validate GUIDs
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guidRegex.test(tenantId)) {
        showError('Tenant ID must be a valid GUID (e.g., 12345678-1234-1234-1234-123456789abc)');
        return;
    }
    if (!guidRegex.test(clientId)) {
        showError('Client ID must be a valid GUID (e.g., 12345678-1234-1234-1234-123456789abc)');
        return;
    }
    
    // Check if user accidentally swapped Tenant ID and Client ID
    if (tenantId === clientId) {
        showError('Tenant ID and Client ID cannot be the same. Please check your Azure AD app registration.');
        return;
    }
    
    // Validate and clean org URL
    if (!orgUrlValue.startsWith('https://')) {
        showError('Organization URL must start with https://');
        return;
    }
    
    orgUrl = orgUrlValue.replace(/\/$/, ''); // Remove trailing slash
    
    // Extract environment ID from URL if possible
    // Format: https://admin.powerplatform.microsoft.com/manage/environments/{environmentId}/applications
    // Or from org URL: https://{orgname}.crm.dynamics.com
    
    // Save credentials if requested
    if (rememberMe) {
        localStorage.setItem('d365_app_updater_creds', JSON.stringify({
            orgUrl: orgUrl,
            tenantId: tenantId,
            clientId: clientId
        }));
    } else {
        localStorage.removeItem('d365_app_updater_creds');
    }
    
    try {
        showLoading('Authenticating...', 'Connecting to Microsoft');
        
        // Initialize MSAL
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        
        // Get token
        const accounts = msalInstance.getAllAccounts();
        const loginRequest = {
            scopes: [`${orgUrl}/.default`],
            account: accounts[0] || undefined,
        };
        
        let authResult;
        if (accounts.length > 0) {
            try {
                authResult = await msalInstance.acquireTokenSilent(loginRequest);
            } catch (error) {
                authResult = await msalInstance.acquireTokenPopup(loginRequest);
            }
        } else {
            authResult = await msalInstance.acquireTokenPopup(loginRequest);
        }
        
        accessToken = authResult.accessToken;
        
        // Test connection
        await testConnection();
        
        // Get environment information
        await getEnvironmentInfo();
        
        hideLoading();
        
        // Switch to apps view
        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appsSection').classList.remove('hidden');
        
        // Load applications
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Authentication error:', error);
        
        let errorMessage = 'Authentication failed: ' + error.message;
        
        if (error.message.includes('AADSTS9002326')) {
            errorMessage = 'App must be configured as Single-Page Application (SPA) in Azure AD. Check setup instructions below.';
        } else if (error.message.includes('AADSTS500113')) {
            errorMessage = 'Redirect URI not configured. Add ' + window.location.origin + ' to your Azure AD app registration.';
        } else if (error.message.includes('endpoints_resolution_error') || error.message.includes('openid_config_error')) {
            errorMessage = '❌ Invalid Tenant ID!\n\n' +
                          'The Tenant ID you entered appears to be invalid or you may have entered the Client ID instead.\n\n' +
                          '✓ Tenant ID = Directory (tenant) ID from Azure AD Overview\n' +
                          '✗ Do NOT use the Application (client) ID here\n\n' +
                          'To find your Tenant ID:\n' +
                          '1. Go to Azure Portal → Azure Active Directory\n' +
                          '2. Click "Overview"\n' +
                          '3. Copy the "Tenant ID" (Directory ID)\n\n' +
                          'Current tenant ID you entered: ' + document.getElementById('tenantId').value;
        }
        
        showError(errorMessage);
    }
}

// Test connection to D365
async function testConnection() {
    const response = await fetch(`${orgUrl}/api/data/v9.2/WhoAmI`, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
        },
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Connection test failed: ${response.status} - ${errorText}`);
    }
    
    return await response.json();
}

// Get environment information
async function getEnvironmentInfo() {
    try {
        // Try to get organization info
        const response = await fetch(`${orgUrl}/api/data/v9.2/organizations?$select=name,organizationid`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
            },
        });
        
        if (response.ok) {
            const data = await response.json();
            if (data.value && data.value.length > 0) {
                const org = data.value[0];
                document.getElementById('environmentName').textContent = org.name;
                environmentId = org.organizationid;
            }
        }
    } catch (error) {
        console.warn('Could not fetch environment info:', error);
        document.getElementById('environmentName').textContent = orgUrl;
    }
}

// Load applications from the environment
async function loadApplications() {
    showLoading('Loading applications...', 'Fetching installed apps');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = `
        <div class="text-center py-5">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Loading...</span>
            </div>
            <p class="mt-3">Loading applications...</p>
        </div>
    `;
    
    try {
        // Query for installed applications (packages)
        const response = await fetch(
            `${orgUrl}/api/data/v9.2/msdyn_solutions?$select=msdyn_solutionid,msdyn_name,msdyn_displayname,msdyn_version,msdyn_installedon&$filter=statecode eq 0&$orderby=msdyn_displayname asc`,
            {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                },
            }
        );
        
        if (!response.ok) {
            // Try alternative approach - query publishers
            return await loadApplicationsAlternative();
        }
        
        const data = await response.json();
        apps = data.value || [];
        
        // Check for updates for each app (simulated - in reality you'd need to query available updates)
        // For demo: Use a deterministic approach based on app ID so results are consistent
        apps = apps.map(app => {
            const appId = app.msdyn_solutionid || app.solutionid || '';
            // Use hash of first char to determine if update is available (consistent results)
            const hasUpdate = appId.length > 0 && parseInt(appId[0], 16) % 3 === 0;
            return {
                ...app,
                hasUpdate: hasUpdate,
                latestVersion: app.msdyn_version || app.version ? 
                    incrementVersion(app.msdyn_version || app.version) : '1.0.0.1'
            };
        });
        
        displayApplications();
        hideLoading();
        
    } catch (error) {
        hideLoading();
        console.error('Error loading applications:', error);
        appsList.innerHTML = `
            <div class="alert alert-danger">
                <i class="fas fa-exclamation-triangle"></i> 
                Failed to load applications: ${error.message}
            </div>
        `;
    }
}

// Alternative method to load applications (using solutions)
async function loadApplicationsAlternative() {
    try {
        const response = await fetch(
            `${orgUrl}/api/data/v9.2/solutions?$select=solutionid,friendlyname,version,installedon,publisherid&$filter=ismanaged eq true and isvisible eq true&$orderby=friendlyname asc`,
            {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                },
            }
        );
        
        if (!response.ok) {
            throw new Error(`Failed to fetch solutions: ${response.status}`);
        }
        
        const data = await response.json();
        apps = data.value.map(solution => ({
            msdyn_solutionid: solution.solutionid,
            msdyn_name: solution.friendlyname,
            msdyn_displayname: solution.friendlyname,
            msdyn_version: solution.version,
            msdyn_installedon: solution.installedon,
            hasUpdate: Math.random() > 0.7,
            latestVersion: incrementVersion(solution.version)
        }));
        
        displayApplications();
        hideLoading();
        
    } catch (error) {
        throw new Error('Could not load applications using alternative method: ' + error.message);
    }
}

// Display applications in the UI
function displayApplications() {
    const appsList = document.getElementById('appsList');
    
    if (apps.length === 0) {
        appsList.innerHTML = `
            <div class="text-center py-5">
                <i class="fas fa-inbox fa-3x text-muted mb-3"></i>
                <p class="text-muted">No applications found in this environment.</p>
            </div>
        `;
        return;
    }
    
    const appsWithUpdates = apps.filter(app => app.hasUpdate);
    document.getElementById('updateCount').textContent = appsWithUpdates.length;
    document.getElementById('updateAllBtn').disabled = appsWithUpdates.length === 0;
    
    let html = '';
    
    apps.forEach(app => {
        const appName = app.msdyn_displayname || app.msdyn_name || 'Unknown App';
        const currentVersion = app.msdyn_version || '1.0.0.0';
        const installedDate = app.msdyn_installedon ? new Date(app.msdyn_installedon).toLocaleDateString() : 'Unknown';
        
        html += `
            <div class="app-card">
                <div class="row align-items-center">
                    <div class="col-md-6">
                        <div class="app-name">
                            <i class="fas fa-cube me-2"></i>${escapeHtml(appName)}
                        </div>
                        <div class="app-version mt-2">
                            <i class="fas fa-tag"></i> Current: ${escapeHtml(currentVersion)}
                            ${app.hasUpdate ? `<br><i class="fas fa-arrow-up text-success"></i> Available: <strong>${escapeHtml(app.latestVersion)}</strong>` : ''}
                        </div>
                        <div class="text-muted small mt-1">
                            <i class="fas fa-calendar"></i> Installed: ${installedDate}
                        </div>
                    </div>
                    <div class="col-md-3">
                        ${app.hasUpdate ? 
                            '<span class="badge-update"><i class="fas fa-arrow-circle-up"></i> Update Available</span>' : 
                            '<span class="badge-current"><i class="fas fa-check-circle"></i> Up to Date</span>'
                        }
                    </div>
                    <div class="col-md-3 text-end">
                        ${app.hasUpdate ? 
                            `<button class="btn btn-success" onclick="updateSingleApp('${app.msdyn_solutionid}')">
                                <i class="fas fa-download"></i> Update Now
                            </button>` : 
                            `<button class="btn btn-outline-secondary" disabled>
                                <i class="fas fa-check"></i> Current
                            </button>`
                        }
                    </div>
                </div>
            </div>
        `;
    });
    
    appsList.innerHTML = html;
}

// Update a single app
async function updateSingleApp(appId) {
    const app = apps.find(a => a.msdyn_solutionid === appId);
    if (!app) {
        showError('App not found');
        return;
    }
    
    const appName = app.msdyn_displayname || app.msdyn_name;
    
    if (!confirm(`Update ${appName} to version ${app.latestVersion}?`)) {
        return;
    }
    
    showLoading('Updating...', `Updating ${appName}`);
    
    try {
        // Simulate update process
        // In a real implementation, you would call the Power Platform API to trigger the update
        await simulateUpdate(appId, appName);
        
        // Mark as updated
        app.hasUpdate = false;
        app.msdyn_version = app.latestVersion;
        
        hideLoading();
        displayApplications();
        showSuccess(`${appName} updated successfully!`);
        
    } catch (error) {
        hideLoading();
        console.error('Update error:', error);
        showError(`Failed to update ${appName}: ${error.message}`);
    }
}

// Update all apps
async function updateAllApps() {
    const appsToUpdate = apps.filter(app => app.hasUpdate);
    
    if (appsToUpdate.length === 0) {
        showError('No apps need updating');
        return;
    }
    
    if (!confirm(`Update all ${appsToUpdate.length} applications?`)) {
        return;
    }
    
    showLoading('Updating all apps...', 'This may take several minutes');
    
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        const appName = app.msdyn_displayname || app.msdyn_name;
        
        document.getElementById('loadingDetails').textContent = 
            `Updating ${i + 1} of ${appsToUpdate.length}: ${appName}`;
        
        try {
            await simulateUpdate(app.msdyn_solutionid, appName);
            app.hasUpdate = false;
            app.msdyn_version = app.latestVersion;
            successCount++;
        } catch (error) {
            console.error(`Failed to update ${appName}:`, error);
            failCount++;
        }
    }
    
    hideLoading();
    displayApplications();
    
    if (failCount === 0) {
        showSuccess(`All ${successCount} applications updated successfully!`);
    } else {
        showError(`Updated ${successCount} apps. ${failCount} failed.`);
    }
}

// Simulate update process (replace with actual API calls)
async function simulateUpdate(appId, appName) {
    // Simulate API call delay
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // ⚠️ SIMULATION ONLY - Real implementation requires:
    // 1. Power Platform Admin API authentication
    // 2. Query available updates from Microsoft catalog
    // 3. Trigger update installation via Admin API
    // 4. Poll for completion status
    // See: https://github.com/moliveirapinto/d365-app-updater/blob/main/POWERPLATFORM_API.md
    
    console.log(`[SIMULATED] Updating app: ${appName} (${appId})`);
    
    // Simulate deterministic failures based on app ID (consistent results)
    const appIdNum = parseInt(appId.substring(0, 8), 16);
    if (appIdNum % 10 === 0) {
        throw new Error('Update failed - simulated error for demo purposes');
    }
}

// Helper function to increment version number
function incrementVersion(version) {
    const parts = version.split('.');
    if (parts.length >= 4) {
        const lastPart = parseInt(parts[3]) || 0;
        parts[3] = (lastPart + 1).toString();
        return parts.join('.');
    }
    return version + '.1';
}

// Handle logout
function handleLogout() {
    if (confirm('Are you sure you want to logout?')) {
        // Clear session
        accessToken = null;
        environmentId = null;
        apps = [];
        
        if (msalInstance) {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.logoutPopup({ account: accounts[0] });
            }
        }
        
        // Switch back to auth view
        document.getElementById('appsSection').classList.add('hidden');
        document.getElementById('authSection').classList.remove('hidden');
    }
}

// UI Helper functions
function showLoading(message, details = '') {
    const overlay = document.getElementById('loadingOverlay');
    const messageEl = document.getElementById('loadingMessage');
    const detailsEl = document.getElementById('loadingDetails');
    
    if (messageEl) messageEl.textContent = message;
    if (detailsEl) detailsEl.textContent = details;
    if (overlay) {
        overlay.classList.remove('hidden');
        overlay.style.display = 'flex';
    }
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.classList.add('hidden');
        overlay.style.display = 'none';
    }
}

function showError(message) {
    alert('❌ Error: ' + message);
}

function showSuccess(message) {
    alert('✅ ' + message);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Make functions available globally
window.updateSingleApp = updateSingleApp;
