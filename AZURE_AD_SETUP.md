# Azure AD App Registration Setup Guide

Follow these steps to create and configure your Azure AD application for the D365 App Updater.

## Step 1: Create the App Registration

1. Navigate to https://portal.azure.com
2. Click on **Azure Active Directory** in the left sidebar
3. Click on **App registrations**
4. Click **+ New registration**

## Step 2: Configure Basic Settings

Fill in the registration form:

- **Name**: `D365 Power Platform App Updater`
- **Supported account types**: 
  - Select "Accounts in this organizational directory only (Single tenant)"
- **Redirect URI**: 
  - Leave blank for now (we'll add it in the next step)

Click **Register**

## Step 3: Configure Authentication

After the app is created:

1. In the left menu, click **Authentication**
2. Click **+ Add a platform**
3. Select **Single-page application**
4. Add your Redirect URI:
   - For local testing: `http://localhost:8000`
   - For production: Your deployed app URL (e.g., `https://yourapp.azurestaticapps.net`)
   - You can add multiple URIs for different environments
5. Under **Implicit grant and hybrid flows**, check:
   - ✅ **Access tokens** (used for implicit flows)
   - ✅ **ID tokens** (used for implicit flows)
6. Click **Configure**
7. Click **Save** at the top

## Step 4: Add API Permissions

⚠️ **CRITICAL**: You must add TWO API permissions for the app to work.

### Add Power Platform API Permission

1. In the left menu, click **API permissions**
2. Click **+ Add a permission**
3. Click **APIs my organization uses**
4. Search for: **"Power Platform API"** or paste `https://api.powerplatform.com`
5. Select **Delegated permissions**
6. Check the box for **user_impersonation**
7. Click **Add permissions**

### Add BAP API Permission

1. Click **+ Add a permission** again
2. Click **APIs my organization uses**
3. Search for: **"BAP"** or paste `https://api.bap.microsoft.com`
4. Select **Delegated permissions**
5. Check the box for **user_impersonation**
6. Click **Add permissions**

### Grant Consent

1. **Important**: Click **Grant admin consent for [Your Org]** at the top (requires admin privileges)
   - If you don't see this button, contact your Azure AD administrator
2. You should now see both APIs with green checkmarks

## Step 5: Get Your Credentials

1. Go to the **Overview** page of your app registration
2. Copy these values (you'll need them in the app):
   - **Application (client) ID**: Example: `12345678-1234-1234-1234-123456789abc`
   - **Directory (tenant) ID**: Example: `87654321-4321-4321-4321-abcdef123456`

## Step 6: Test the Configuration

1. Open your D365 App Updater application
2. Enter:
   - **Organization URL**: Your D365 URL (e.g., `https://yourorg.crm.dynamics.com`)
   - **Tenant ID**: The Directory (tenant) ID from step 5
   - **Client ID**: The Application (client) ID from step 5
3. Click **Connect to Power Platform**
4. You should see a Microsoft login popup
5. Sign in with your credentials
6. Grant consent if prompted

## Common Configuration Issues

### Issue: "AADSTS9002326: Cross-origin token redemption"

**Problem**: App is configured as "Web" instead of "Single-page application"

**Solution**:
1. Go to Authentication settings
2. If you see a "Web" platform, remove it
3. Add a "Single-page application" platform instead
4. Add your redirect URI to the SPA platform

### Issue: "AADSTS500113: No reply address"

**Problem**: Redirect URI doesn't match or isn't configured

**Solution**:
1. Make sure your redirect URI exactly matches the URL you're accessing the app from
2. Include protocol (http:// or https://)
3. Port numbers must match if using localhost
4. No trailing slashes

### Issue: "AADSTS65001: User or admin has not consented"

**Problem**: API permissions not granted

**Solution**:
1. Go to API permissions
2. Click "Grant admin consent for [Your Org]"
3. If you can't do this, contact your Azure AD administrator

### Issue: "AADSTS650057: Invalid resource"

**Error message**: "The client has requested access to a resource which is not listed in the requested permissions"

**Problem**: You're missing the Power Platform API and/or BAP API permissions

**Solution**:
1. Go to **API permissions**
2. Verify you have BOTH:
   - ✅ Power Platform API (`https://api.powerplatform.com`)
   - ✅ BAP API (`https://api.bap.microsoft.com`)
3. If either is missing, click **+ Add a permission** → **APIs my organization uses** → search for the API → add **user_impersonation** permission
4. Click **Grant admin consent**

### Issue: "Access token validation failure"

**Problem**: Wrong tenant ID or missing permissions

**Solution**:
1. Verify you're using the correct Tenant ID
2. Ensure Power Platform API and BAP API permissions are added and consented
3. Make sure you're signing in with an account that has access to the D365 environment

## Security Best Practices

1. **Never expose the Client Secret** (not needed for SPA apps anyway)
2. **Use separate app registrations** for development and production
3. **Regularly review permissions** and remove unused ones
4. **Enable logging** in Azure AD to monitor authentication attempts
5. **Use Conditional Access policies** if required by your organization

## Multi-Environment Setup

If you need to support multiple environments (dev, test, prod):

### Option 1: Multiple Redirect URIs (Recommended)
- Add all environment URLs to the same app registration
- Example:
  - `http://localhost:8000`
  - `https://dev-appupdater.azurestaticapps.net`
  - `https://appupdater.azurestaticapps.net`

### Option 2: Separate App Registrations
- Create a different app registration for each environment
- More secure but requires managing multiple client IDs

## Testing Checklist

Before deploying to users:

- [ ] App registration created
- [ ] Single-page application platform configured
- [ ] Correct redirect URI(s) added
- [ ] Access tokens and ID tokens enabled
- [ ] Power Platform API permission added (`https://api.powerplatform.com`)
- [ ] BAP API permission added (`https://api.bap.microsoft.com`)
- [ ] Admin consent granted for both APIs
- [ ] Test login successful
- [ ] Can access D365 environment data
- [ ] Apps list loads correctly

## Additional Resources

- [Microsoft Identity Platform Documentation](https://docs.microsoft.com/azure/active-directory/develop/)
- [Register an application](https://docs.microsoft.com/azure/active-directory/develop/quickstart-register-app)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Dynamics 365 Web API](https://docs.microsoft.com/dynamics365/customer-engagement/web-api/)

## Need Help?

If you encounter issues:

1. Check the browser console for detailed error messages
2. Review Azure AD sign-in logs (Azure AD → Sign-in logs)
3. Verify all permissions are granted
4. Test with a simple WhoAmI API call first
5. Contact your Azure AD administrator if you can't grant consent

---

Last updated: January 2026
