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
    const authForm = document.getElementById('authForm');
    if (authForm) {
        authForm.addEventListener('submit', handleAuthentication);
        console.log('Auth form listener attached');
    } else {
        console.error('Auth form not found!');
    }
    
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
    
    console.log('Authentication started');
    
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
    
    orgUrl = orgUrlValue.replace(/\/$/, '');
    
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
            errorMessage = 'Invalid Tenant ID! The Tenant ID you entered appears to be invalid. Please check your Azure AD settings.';
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
                console.log('Environment ID:', environmentId);
                return;
            }
        }
    } catch (error) {
        console.warn('Could not fetch environment info:', error);
    }
    
    document.getElementById('environmentName').textContent = orgUrl;
}

// Load applications from the environment using Dataverse API
async function loadApplications() {
    showLoading('Loading applications...', 'Fetching installed solutions');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div><p class="mt-3">Loading applications...</p></div>';
    
    try {
        // Query solutions from Dataverse (these are the installed apps/packages)
        showLoading('Loading applications...', 'Querying installed solutions');
        
        // Get all solutions that are managed (typically the apps that can be updated)
        const url = `${orgUrl}/api/data/v9.2/solutions?$filter=ismanaged eq true&$select=solutionid,uniquename,friendlyname,version,installedon,publisherid,ismanaged&$expand=publisherid($select=friendlyname)&$orderby=friendlyname`;
        
        console.log('Fetching solutions from:', url);
        
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
                'Prefer': 'odata.include-annotations="*"'
            },
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Error:', response.status, errorText);
            throw new Error(`Failed to fetch solutions: ${response.status}`);
        }
        
        const data = await response.json();
        console.log('Solutions data:', data);
        
        apps = (data.value || []).map(solution => ({
            id: solution.solutionid,
            uniqueName: solution.uniquename,
            name: solution.friendlyname || solution.uniquename,
            version: solution.version || '1.0.0.0',
            installedOn: solution.installedon,
            publisher: solution.publisherid ? solution.publisherid.friendlyname : 'Unknown',
            isManaged: solution.ismanaged,
            hasUpdate: false,
            latestVersion: solution.version || '1.0.0.0',
            updateAvailable: null
        }));
        
        // Check for available updates via staging
        showLoading('Checking for updates...', 'Checking ' + apps.length + ' solutions');
        await checkForUpdates();
        
        displayApplications();
        hideLoading();
        
    } catch (error) {
        hideLoading();
        console.error('Error loading applications:', error);
        appsList.innerHTML = '<div class="alert alert-danger"><i class="fas fa-exclamation-triangle"></i> <strong>Failed to load applications</strong><br>' + error.message + '</div>';
    }
}

// Check for available updates using solution staging
async function checkForUpdates() {
    try {
        // Query staged solutions (solutions available to upgrade)
        const stagingUrl = `${orgUrl}/api/data/v9.2/stagesolutionuploads?$select=solutionuniquename,solutionversion,name`;
        
        const response = await fetch(stagingUrl, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
            },
        });
        
        if (response.ok) {
            const stagingData = await response.json();
            console.log('Staged solutions:', stagingData);
            
            if (stagingData.value && stagingData.value.length > 0) {
                for (let i = 0; i < apps.length; i++) {
                    const app = apps[i];
                    // Find if there's a staged version of this solution
                    const stagedSolution = stagingData.value.find(function(s) {
                        return s.solutionuniquename === app.uniqueName;
                    });
                    
                    if (stagedSolution && stagedSolution.solutionversion !== app.version) {
                        app.hasUpdate = true;
                        app.latestVersion = stagedSolution.solutionversion;
                        app.updateAvailable = stagedSolution;
                    }
                }
            }
        }
    } catch (error) {
        console.warn('Could not check for staged updates:', error);
    }
    
    // Also check solution history for any pending upgrades
    try {
        const historyUrl = `${orgUrl}/api/data/v9.2/msdyn_solutionhistories?$filter=msdyn_status eq 0&$select=msdyn_solutionid,msdyn_solutionversion,msdyn_name&$top=100`;
        
        const historyResponse = await fetch(historyUrl, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
            },
        });
        
        if (historyResponse.ok) {
            const historyData = await historyResponse.json();
            console.log('Solution history:', historyData);
        }
    } catch (error) {
        console.warn('Could not check solution history:', error);
    }
}

// Display applications in the UI
function displayApplications() {
    const appsList = document.getElementById('appsList');
    
    if (apps.length === 0) {
        appsList.innerHTML = '<div class="text-center py-5"><i class="fas fa-inbox fa-3x text-muted mb-3"></i><p class="text-muted">No applications found in this environment.</p></div>';
        return;
    }
    
    const appsWithUpdates = apps.filter(function(app) { return app.hasUpdate; });
    document.getElementById('updateCount').textContent = appsWithUpdates.length;
    document.getElementById('updateAllBtn').disabled = appsWithUpdates.length === 0;
    
    let html = '';
    
    for (let i = 0; i < apps.length; i++) {
        const app = apps[i];
        const appName = app.name || 'Unknown App';
        const currentVersion = app.version || '1.0.0.0';
        const installedDate = app.installedOn ? new Date(app.installedOn).toLocaleDateString() : 'Unknown';
        
        html += '<div class="app-card">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-6">';
        html += '<div class="app-name"><i class="fas fa-cube me-2"></i>' + escapeHtml(appName) + '</div>';
        html += '<div class="app-version mt-2"><i class="fas fa-tag"></i> Current: ' + escapeHtml(currentVersion);
        if (app.hasUpdate) {
            html += '<br><i class="fas fa-arrow-up text-success"></i> Available: <strong>' + escapeHtml(app.latestVersion) + '</strong>';
        }
        html += '</div>';
        html += '<div class="text-muted small mt-1"><i class="fas fa-calendar"></i> Installed: ' + installedDate + '</div>';
        html += '</div>';
        html += '<div class="col-md-3">';
        if (app.hasUpdate) {
            html += '<span class="badge-update"><i class="fas fa-arrow-circle-up"></i> Update Available</span>';
        } else {
            html += '<span class="badge-current"><i class="fas fa-check-circle"></i> Up to Date</span>';
        }
        html += '</div>';
        html += '<div class="col-md-3 text-end">';
        if (app.hasUpdate) {
            html += '<button class="btn btn-success" onclick="updateSingleApp(\'' + app.id + '\')"><i class="fas fa-download"></i> Update Now</button>';
        } else {
            html += '<button class="btn btn-outline-secondary" disabled><i class="fas fa-check"></i> Current</button>';
        }
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    
    appsList.innerHTML = html;
}

// Update a single solution
async function updateSingleApp(appId) {
    const app = apps.find(function(a) { return a.id === appId; });
    if (!app) {
        showError('Solution not found');
        return;
    }
    
    const appName = app.name;
    
    if (!confirm('Upgrade solution "' + appName + '"?\n\nNote: Solution upgrades in Dataverse require importing a new version of the solution. If you have a staged solution ready, it will be applied.\n\nFor Microsoft first-party apps, use the Power Platform Admin Center.')) {
        return;
    }
    
    showLoading('Processing...', 'Checking solution ' + appName);
    
    try {
        // For managed solutions, we need to trigger the upgrade via solution import
        // This requires the staged solution to be available
        if (app.updateAvailable) {
            const stagingSolutionId = app.updateAvailable.stagesolutionuploadid;
            
            // Import the staged solution
            const importUrl = `${orgUrl}/api/data/v9.2/ImportSolutionAsync`;
            
            const response = await fetch(importUrl, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                },
                body: JSON.stringify({
                    OverwriteUnmanagedCustomizations: false,
                    PublishWorkflows: true,
                    StageSolutionUploadId: stagingSolutionId
                })
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error('Import failed:', response.status, errorText);
                throw new Error('Solution import failed: ' + response.status);
            }
            
            const result = await response.json();
            console.log('Import started:', result);
            
            // Mark as updated
            app.hasUpdate = false;
            app.version = app.latestVersion;
            
            hideLoading();
            displayApplications();
            showSuccess(appName + ' upgrade initiated! Check the Power Platform Admin Center for progress.');
        } else {
            hideLoading();
            showError('No staged update available for ' + appName + '. Please upload the new solution version first, or use the Power Platform Admin Center for Microsoft apps.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('Update error:', error);
        showError('Failed to update ' + appName + ': ' + error.message);
    }
}

// Poll for import job completion status
async function pollImportStatus(asyncOperationId, appName) {
    let attempts = 0;
    const maxAttempts = 60;
    
    while (attempts < maxAttempts) {
        try {
            const statusUrl = `${orgUrl}/api/data/v9.2/asyncoperations(${asyncOperationId})?$select=statuscode,statecode,message`;
            
            const response = await fetch(statusUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                },
            });
            
            if (response.ok) {
                const operation = await response.json();
                const statusCode = operation.statuscode;
                
                // StatusCode: 30 = Succeeded, 31 = Failed, 32 = Cancelled
                if (statusCode === 30) {
                    return true;
                } else if (statusCode === 31 || statusCode === 32) {
                    throw new Error('Import failed: ' + (operation.message || 'Unknown error'));
                }
                
                document.getElementById('loadingDetails').textContent = 'Installing ' + appName + '... (' + (attempts * 5) + 's elapsed)';
            }
        } catch (error) {
            console.warn('Error polling status:', error);
        }
        
        await new Promise(function(resolve) { setTimeout(resolve, 5000); });
        attempts++;
    }
    
    throw new Error('Import timeout - operation took too long');
}

// Update all apps with staged updates
async function updateAllApps() {
    const appsToUpdate = apps.filter(function(app) { return app.hasUpdate && app.updateAvailable; });
    
    if (appsToUpdate.length === 0) {
        // Show info about how to get updates
        alert('No staged solution updates available.\n\nTo update Microsoft first-party apps (like Sales Hub, Customer Service Hub), please use the Power Platform Admin Center:\n\n1. Go to admin.powerplatform.microsoft.com\n2. Select your environment\n3. Go to Resources > Dynamics 365 apps\n4. Click "Manage" to install updates');
        return;
    }
    
    if (!confirm('Import ' + appsToUpdate.length + ' staged solution updates?\n\nThis will upgrade the solutions in your environment.')) {
        return;
    }
    
    showLoading('Updating solutions...', 'This may take several minutes');
    
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < appsToUpdate.length; i++) {
        const app = appsToUpdate[i];
        const appName = app.name;
        
        document.getElementById('loadingDetails').textContent = 'Updating ' + (i + 1) + ' of ' + appsToUpdate.length + ': ' + appName;
        
        try {
            if (app.updateAvailable && app.updateAvailable.stagesolutionuploadid) {
                const importUrl = `${orgUrl}/api/data/v9.2/ImportSolutionAsync`;
                
                const response = await fetch(importUrl, {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0',
                    },
                    body: JSON.stringify({
                        OverwriteUnmanagedCustomizations: false,
                        PublishWorkflows: true,
                        StageSolutionUploadId: app.updateAvailable.stagesolutionuploadid
                    })
                });
                
                if (!response.ok) {
                    throw new Error('HTTP ' + response.status);
                }
                
                const result = await response.json();
                
                if (result.AsyncOperationId) {
                    await pollImportStatus(result.AsyncOperationId, appName);
                }
                
                app.hasUpdate = false;
                app.version = app.latestVersion;
                successCount++;
            } else {
                throw new Error('No staged solution available');
            }
            
        } catch (error) {
            console.error('Failed to update ' + appName + ':', error);
            failCount++;
        }
    }
    
    hideLoading();
    displayApplications();
    
    if (failCount === 0) {
        showSuccess('All ' + successCount + ' solutions updated successfully!');
    } else {
        showError('Updated ' + successCount + ' solutions. ' + failCount + ' failed.');
    }
}

// Handle logout
function handleLogout() {
    if (confirm('Are you sure you want to logout?')) {
        accessToken = null;
        environmentId = null;
        apps = [];
        
        if (msalInstance) {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.logoutPopup({ account: accounts[0] });
            }
        }
        
        document.getElementById('appsSection').classList.add('hidden');
        document.getElementById('authSection').classList.remove('hidden');
    }
}

// UI Helper functions
function showLoading(message, details) {
    const overlay = document.getElementById('loadingOverlay');
    const messageEl = document.getElementById('loadingMessage');
    const detailsEl = document.getElementById('loadingDetails');
    
    if (messageEl) messageEl.textContent = message;
    if (detailsEl) detailsEl.textContent = details || '';
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
    hideLoading();
    alert('Error: ' + message);
}

function showSuccess(message) {
    alert(message);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Make functions available globally
window.updateSingleApp = updateSingleApp;
