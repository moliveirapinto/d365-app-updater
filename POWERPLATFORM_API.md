# Power Platform API Implementation Notes

This document contains notes and code samples for implementing the actual Power Platform app update functionality.

## Current Status

The current implementation uses **simulated/mock data** for:
- Available updates detection
- Update installation process
- Update status polling

## Required Power Platform APIs

To make this tool fully functional, you'll need to implement calls to the Power Platform Admin APIs:

### 1. List Installed Applications

**Endpoint**: 
```
GET https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentId}/applicationPackages
```

**Authentication**: Bearer token with scope `https://api.bap.microsoft.com/.default`

**Response**:
```json
{
  "value": [
    {
      "id": "guid",
      "name": "Application Name",
      "displayName": "Display Name",
      "version": "1.0.0.0",
      "installedOn": "2025-01-01T00:00:00Z",
      "publisher": "Microsoft",
      "state": "Installed"
    }
  ]
}
```

### 2. Check for Available Updates

**Endpoint**:
```
GET https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/applicationPackages/{packageId}/versions
```

**Response**:
```json
{
  "value": [
    {
      "version": "1.0.0.1",
      "releaseDate": "2025-01-15T00:00:00Z",
      "isLatest": true,
      "releaseNotes": "Bug fixes and improvements"
    }
  ]
}
```

### 3. Trigger Application Update

**Endpoint**:
```
POST https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{environmentId}/applicationPackages/{packageId}/install
```

**Body**:
```json
{
  "version": "1.0.0.1"
}
```

**Response**:
```json
{
  "operationId": "guid",
  "status": "Running"
}
```

### 4. Check Update Status

**Endpoint**:
```
GET https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/operations/{operationId}
```

**Response**:
```json
{
  "id": "guid",
  "status": "Succeeded|Running|Failed",
  "createdDateTime": "2025-01-01T00:00:00Z",
  "completedDateTime": "2025-01-01T00:05:00Z",
  "error": null
}
```

## Implementation Steps

### Step 1: Update Authentication Scope

The current implementation uses Dynamics 365 scope. You'll also need Power Platform Admin API scope:

```javascript
const loginRequest = {
    scopes: [
        `${orgUrl}/.default`,
        'https://api.bap.microsoft.com/.default'
    ],
    account: accounts[0] || undefined,
};
```

### Step 2: Get Environment ID

You need to extract or obtain the environment ID. Options:

**Option A: From Organization URL**
```javascript
// Extract from admin URL if user provides it
// https://admin.powerplatform.microsoft.com/manage/environments/{envId}/applications
function extractEnvironmentId(url) {
    const match = url.match(/environments\/([a-f0-9-]+)/i);
    return match ? match[1] : null;
}
```

**Option B: Query from Dynamics**
```javascript
async function getEnvironmentId() {
    const response = await fetch(`${orgUrl}/api/data/v9.2/organizations?$select=organizationid`, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
        },
    });
    const data = await response.json();
    return data.value[0].organizationid;
}
```

### Step 3: Implement Real loadApplications()

Replace the current `loadApplications()` function:

```javascript
async function loadApplications() {
    try {
        // Get BAP token (Power Platform Admin API)
        const bapToken = await getBAPToken();
        
        // Get installed apps
        const response = await fetch(
            `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentId}/applicationPackages?api-version=2021-04-01`,
            {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${bapToken}`,
                    'Accept': 'application/json',
                },
            }
        );
        
        if (!response.ok) {
            throw new Error(`Failed to fetch applications: ${response.status}`);
        }
        
        const data = await response.json();
        
        // For each app, check for available updates
        apps = await Promise.all(data.value.map(async (app) => {
            const latestVersion = await getLatestVersion(app.id, bapToken);
            return {
                ...app,
                hasUpdate: latestVersion && latestVersion !== app.version,
                latestVersion: latestVersion || app.version
            };
        }));
        
        displayApplications();
        
    } catch (error) {
        console.error('Error loading applications:', error);
        throw error;
    }
}
```

### Step 4: Implement Real updateSingleApp()

```javascript
async function updateSingleApp(appId) {
    const app = apps.find(a => a.id === appId);
    if (!app) return;
    
    try {
        showLoading('Updating...', `Installing ${app.displayName}`);
        
        const bapToken = await getBAPToken();
        
        // Trigger update
        const response = await fetch(
            `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentId}/applicationPackages/${appId}/install?api-version=2021-04-01`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${bapToken}`,
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    version: app.latestVersion
                })
            }
        );
        
        if (!response.ok) {
            throw new Error(`Update failed: ${response.status}`);
        }
        
        const operation = await response.json();
        
        // Poll for completion
        await pollUpdateStatus(operation.operationId, bapToken);
        
        // Refresh app list
        await loadApplications();
        
        hideLoading();
        showSuccess(`${app.displayName} updated successfully!`);
        
    } catch (error) {
        hideLoading();
        showError(`Failed to update: ${error.message}`);
    }
}
```

### Step 5: Implement Status Polling

```javascript
async function pollUpdateStatus(operationId, bapToken, maxAttempts = 60) {
    let attempts = 0;
    
    while (attempts < maxAttempts) {
        const response = await fetch(
            `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/operations/${operationId}?api-version=2021-04-01`,
            {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${bapToken}`,
                    'Accept': 'application/json',
                },
            }
        );
        
        const operation = await response.json();
        
        if (operation.status === 'Succeeded') {
            return true;
        } else if (operation.status === 'Failed') {
            throw new Error(operation.error?.message || 'Update failed');
        }
        
        // Wait 5 seconds before polling again
        await new Promise(resolve => setTimeout(resolve, 5000));
        attempts++;
        
        // Update progress
        document.getElementById('loadingDetails').textContent = 
            `Installation in progress... (${attempts * 5}s)`;
    }
    
    throw new Error('Update timeout - operation took too long');
}
```

### Step 6: Get BAP Token

```javascript
async function getBAPToken() {
    if (!msalInstance) {
        throw new Error('Not authenticated');
    }
    
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        throw new Error('No account found');
    }
    
    const tokenRequest = {
        scopes: ['https://api.bap.microsoft.com/.default'],
        account: accounts[0],
    };
    
    try {
        const authResult = await msalInstance.acquireTokenSilent(tokenRequest);
        return authResult.accessToken;
    } catch (error) {
        // Require user interaction
        const authResult = await msalInstance.acquireTokenPopup(tokenRequest);
        return authResult.accessToken;
    }
}
```

## Additional Azure AD Configuration

To use the Power Platform Admin API, you need additional permissions:

1. Go to Azure AD App Registration → API permissions
2. Click "Add a permission"
3. Select "APIs my organization uses"
4. Search for "PowerApps-Advisor" or "Dynamics 365 Business Central"
5. Or manually add the API: `https://api.bap.microsoft.com`
6. Select delegated permissions
7. Grant admin consent

## Testing

Before implementing in production:

1. **Test in a sandbox environment**
2. **Verify permissions** work correctly
3. **Handle rate limiting** (implement exponential backoff)
4. **Add error handling** for various failure scenarios
5. **Log all operations** for debugging

## Error Handling

Common errors to handle:

- **403 Forbidden**: User doesn't have admin permissions
- **404 Not Found**: Environment or app not found
- **429 Too Many Requests**: Rate limiting (implement backoff)
- **500 Server Error**: Retry with exponential backoff
- **Timeout**: Update taking too long

## Rate Limiting

The Power Platform APIs have rate limits. Implement:

```javascript
async function apiCallWithRetry(url, options, maxRetries = 3) {
    for (let i = 0; i < maxRetries; i++) {
        const response = await fetch(url, options);
        
        if (response.status === 429) {
            const retryAfter = response.headers.get('Retry-After') || 60;
            await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
            continue;
        }
        
        return response;
    }
    
    throw new Error('Max retries exceeded');
}
```

## Security Considerations

1. **Never expose tokens** in console logs or error messages
2. **Validate environment access** before allowing operations
3. **Implement audit logging** for all update operations
4. **Consider implementing approval workflows** for production updates
5. **Add confirmation dialogs** for bulk operations

## References

- [Power Platform Admin API Documentation](https://docs.microsoft.com/power-platform/admin/programmability-authentication)
- [Power Platform REST API Reference](https://docs.microsoft.com/rest/api/power-platform/)
- [Environment Management](https://docs.microsoft.com/power-platform/admin/api/reference/environments)
- [Application Lifecycle Management](https://docs.microsoft.com/power-platform/alm/)

---

**Note**: This is a guide for implementing the real functionality. The actual API endpoints and request/response formats should be verified with the latest Power Platform API documentation.
