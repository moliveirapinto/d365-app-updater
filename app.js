// Global variables
let msalInstance = null;
let accessToken = null;
let environmentId = null;
let environmentName = null;
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
        console.error('MSAL library not loaded');
        alert('Error: MSAL library failed to load. Please check your internet connection and refresh the page.');
        return;
    }
    
    console.log('MSAL library loaded successfully');
    
    const redirectUriElement = document.getElementById('redirectUri');
    if (redirectUriElement) {
        redirectUriElement.textContent = window.location.origin;
    }
    
    try {
        loadSavedCredentials();
    } catch (error) {
        console.error('Error loading saved credentials:', error);
    }
    
    const authForm = document.getElementById('authForm');
    if (authForm) {
        authForm.addEventListener('submit', handleAuthentication);
    }
    
    document.getElementById('logoutBtn').addEventListener('click', handleLogout);
    document.getElementById('refreshAppsBtn').addEventListener('click', loadApplications);
    
    console.log('App initialized successfully');
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
        showError('Tenant ID must be a valid GUID');
        return;
    }
    if (!guidRegex.test(clientId)) {
        showError('Client ID must be a valid GUID');
        return;
    }
    if (tenantId === clientId) {
        showError('Tenant ID and Client ID cannot be the same');
        return;
    }
    if (!orgUrlValue.startsWith('https://')) {
        showError('Organization URL must start with https://');
        return;
    }
    
    orgUrl = orgUrlValue.replace(/\/$/, '');
    
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
        
        const msalConfig = createMsalConfig(tenantId, clientId);
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        
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
        
        await testConnection();
        await getEnvironmentInfo();
        
        // Set Admin Center link
        updateAdminCenterLink();
        
        hideLoading();
        
        document.getElementById('authSection').classList.add('hidden');
        document.getElementById('appsSection').classList.remove('hidden');
        
        await loadApplications();
        
    } catch (error) {
        hideLoading();
        console.error('Authentication error:', error);
        
        let errorMessage = 'Authentication failed: ' + error.message;
        
        if (error.message.includes('AADSTS9002326')) {
            errorMessage = 'App must be configured as Single-Page Application (SPA) in Azure AD.';
        } else if (error.message.includes('AADSTS500113')) {
            errorMessage = 'Redirect URI not configured. Add ' + window.location.origin + ' to your Azure AD app registration.';
        } else if (error.message.includes('endpoints_resolution_error') || error.message.includes('openid_config_error')) {
            errorMessage = 'Invalid Tenant ID! Please check your Azure AD settings.';
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
        throw new Error(`Connection test failed: ${response.status}`);
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
                environmentName = org.name;
                environmentId = org.organizationid;
                document.getElementById('environmentName').textContent = environmentName;
                console.log('Environment ID:', environmentId);
                return;
            }
        }
    } catch (error) {
        console.warn('Could not fetch environment info:', error);
    }
    
    document.getElementById('environmentName').textContent = orgUrl;
}

// Update Admin Center link
function updateAdminCenterLink() {
    const adminCenterBtn = document.getElementById('adminCenterBtn');
    if (adminCenterBtn && environmentId) {
        // Link to Power Platform Admin Center for the specific environment
        adminCenterBtn.href = `https://admin.powerplatform.microsoft.com/environments/${environmentId}/applications`;
    }
}

// Load applications from Dataverse
async function loadApplications() {
    showLoading('Loading solutions...', 'Fetching installed solutions from Dataverse');
    
    const appsList = document.getElementById('appsList');
    appsList.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div><p class="mt-3">Loading...</p></div>';
    
    try {
        // Get managed solutions
        const solutionsUrl = `${orgUrl}/api/data/v9.2/solutions?$filter=ismanaged eq true&$select=solutionid,uniquename,friendlyname,version,installedon,publisherid&$expand=publisherid($select=friendlyname)&$orderby=friendlyname`;
        
        const response = await fetch(solutionsUrl, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
            },
        });
        
        if (!response.ok) {
            throw new Error(`Failed to fetch solutions: ${response.status}`);
        }
        
        const data = await response.json();
        console.log('Solutions:', data.value?.length || 0);
        
        apps = (data.value || []).map(solution => ({
            id: solution.solutionid,
            uniqueName: solution.uniquename,
            name: solution.friendlyname || solution.uniquename,
            version: solution.version || '1.0.0.0',
            installedOn: solution.installedon,
            publisher: solution.publisherid ? solution.publisherid.friendlyname : 'Unknown'
        }));
        
        // Filter to show main/important solutions (not internal platform components)
        const importantApps = apps.filter(app => {
            const name = app.name.toLowerCase();
            const uniqueName = app.uniqueName.toLowerCase();
            
            // Filter out internal platform solutions
            const isInternal = 
                uniqueName.startsWith('msft_') ||
                uniqueName.startsWith('apiextension') ||
                uniqueName.endsWith('_anchor') ||
                uniqueName.includes('infrastructure') ||
                uniqueName.includes('infra_') ||
                name.includes('anchor') ||
                name.includes('infrastructure');
            
            // Keep Microsoft Dynamics and important apps
            const isImportant = 
                uniqueName.startsWith('msdyn') ||
                uniqueName.startsWith('msdynce') ||
                uniqueName.startsWith('dynamics') ||
                uniqueName.startsWith('omnichannelprime') ||
                uniqueName.includes('sales') ||
                uniqueName.includes('service') ||
                uniqueName.includes('marketing') ||
                uniqueName.includes('customerinsights') ||
                name.toLowerCase().includes('dynamics') ||
                name.toLowerCase().includes('sales') ||
                name.toLowerCase().includes('service') ||
                name.toLowerCase().includes('marketing');
            
            return isImportant && !isInternal;
        });
        
        displayApplications(importantApps.length > 0 ? importantApps : apps.slice(0, 100));
        hideLoading();
        
    } catch (error) {
        hideLoading();
        console.error('Error loading applications:', error);
        appsList.innerHTML = '<div class="alert alert-danger"><i class="fas fa-exclamation-triangle"></i> Failed to load: ' + error.message + '</div>';
    }
}

// Display applications
function displayApplications(appsToShow) {
    const appsList = document.getElementById('appsList');
    
    if (!appsToShow || appsToShow.length === 0) {
        appsList.innerHTML = '<div class="text-center py-5"><i class="fas fa-inbox fa-3x text-muted mb-3"></i><p class="text-muted">No solutions found.</p></div>';
        return;
    }
    
    document.getElementById('updateCount').textContent = appsToShow.length;
    
    let html = '';
    
    for (let i = 0; i < appsToShow.length; i++) {
        const app = appsToShow[i];
        const appName = app.name || 'Unknown';
        const version = app.version || '1.0.0.0';
        const installedDate = app.installedOn ? new Date(app.installedOn).toLocaleDateString() : 'N/A';
        const publisher = app.publisher || 'Unknown';
        
        // Determine if it's a Microsoft app
        const isMicrosoft = publisher.toLowerCase().includes('microsoft') || publisher.toLowerCase().includes('dynamics');
        
        html += '<div class="app-card">';
        html += '<div class="row align-items-center">';
        html += '<div class="col-md-7">';
        html += '<div class="app-name">';
        html += '<i class="fas fa-cube me-2" style="color: ' + (isMicrosoft ? '#0078d4' : '#6c757d') + ';"></i>';
        html += escapeHtml(appName);
        html += '</div>';
        html += '<div class="app-version mt-2">';
        html += '<i class="fas fa-tag"></i> Version: <strong>' + escapeHtml(version) + '</strong>';
        html += '</div>';
        html += '<div class="text-muted small mt-1">';
        html += '<i class="fas fa-building"></i> ' + escapeHtml(publisher);
        html += ' &nbsp;|&nbsp; <i class="fas fa-calendar"></i> Installed: ' + installedDate;
        html += '</div>';
        html += '</div>';
        html += '<div class="col-md-3 text-center">';
        if (isMicrosoft) {
            html += '<span class="badge bg-primary"><i class="fas fa-microsoft me-1"></i> Microsoft</span>';
        } else {
            html += '<span class="badge bg-secondary"><i class="fas fa-box me-1"></i> Partner</span>';
        }
        html += '</div>';
        html += '<div class="col-md-2 text-end">';
        html += '<span class="badge-current"><i class="fas fa-check-circle"></i> Installed</span>';
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    
    // Add note at bottom
    html += '<div class="alert alert-warning mt-3">';
    html += '<i class="fas fa-lightbulb"></i> <strong>Tip:</strong> To check for updates to these solutions, click ';
    html += '"<strong>Open Admin Center</strong>" above. The Power Platform Admin Center shows available updates for ';
    html += 'Dynamics 365 apps that can be installed with one click.';
    html += '</div>';
    
    appsList.innerHTML = html;
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
                msalInstance.logoutPopup({ account: accounts[0] }).catch(console.error);
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

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
