# Azure Function for Scheduled Auto-Updates

This Azure Function runs on a timer and executes scheduled app updates automatically.

## Architecture Overview

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   D365 App      │────▶│    Supabase     │◀────│ Azure Function  │
│   Updater UI    │     │   (schedules)   │     │  (Timer Trigger)│
└─────────────────┘     └─────────────────┘     └────────┬────────┘
                                                         │
                                                         ▼
                                                ┌─────────────────┐
                                                │  Power Platform │
                                                │      APIs       │
                                                └─────────────────┘
```

## Prerequisites

1. **Azure Subscription** with permission to create Function Apps
2. **Azure AD App Registration** (Service Principal) with:
   - Application permissions (not delegated) for Power Platform APIs
   - Admin consent granted
3. **Supabase** project with the `update_schedules` table
4. **Node.js 18+** for local development

## Setup Instructions

### 1. Create Service Principal for Automation

The Azure Function needs its own identity to update apps on behalf of users. This requires a **service principal** with application permissions.

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Click **New registration**
   - Name: `D365 App Updater - Automation Service`
   - Supported account types: **Accounts in this organizational directory only**
3. After creation, go to **Certificates & secrets** → **New client secret**
   - Copy the secret value (you won't see it again)
4. Go to **API permissions** → **Add a permission**:
   - **Dynamics CRM** → **Application permissions** → `user_impersonation` (if available)
   - **Power Platform API** → **Application permissions** (check available permissions)
   - **Microsoft Graph** → **Application permissions** → `User.Read.All` (to resolve user info)
5. Click **Grant admin consent for [Org]**
6. Copy these values:
   - **Application (client) ID**
   - **Directory (tenant) ID**
   - **Client Secret**

> ⚠️ **Important**: Application permissions for Power Platform/Dynamics may be limited. You may need to configure the service principal as an **Application User** in each Power Platform environment.

### 2. Configure Application User in Power Platform

For each environment where you want auto-updates:

1. Go to [Power Platform Admin Center](https://admin.powerplatform.microsoft.com)
2. Select your environment → **Settings** → **Users + permissions** → **Application users**
3. Click **New app user**
4. Select your service principal
5. Assign **System Administrator** security role (or a custom role with app management permissions)

### 3. Create Azure Function App

```bash
# Install Azure Functions Core Tools
npm install -g azure-functions-core-tools@4

# Login to Azure
az login

# Create a resource group
az group create --name rg-d365-app-updater --location eastus

# Create a storage account (required for Functions)
az storage account create \
  --name std365appupdater \
  --resource-group rg-d365-app-updater \
  --sku Standard_LRS

# Create the Function App
az functionapp create \
  --name fn-d365-app-updater \
  --resource-group rg-d365-app-updater \
  --storage-account std365appupdater \
  --consumption-plan-location eastus \
  --runtime node \
  --runtime-version 18 \
  --functions-version 4
```

### 4. Configure Function App Settings

```bash
az functionapp config appsettings set \
  --name fn-d365-app-updater \
  --resource-group rg-d365-app-updater \
  --settings \
    "AZURE_TENANT_ID=<your-tenant-id>" \
    "AZURE_CLIENT_ID=<your-client-id>" \
    "AZURE_CLIENT_SECRET=<your-client-secret>" \
    "SUPABASE_URL=https://your-project.supabase.co" \
    "SUPABASE_KEY=<your-service-role-key>"
```

> 💡 Use the **service_role** key (not anon) for the Azure Function so it can update `last_run_at` and `last_run_status`.

### 5. Deploy the Function

```bash
cd azure-function
func azure functionapp publish fn-d365-app-updater
```

## Local Development

1. Copy `local.settings.json.example` to `local.settings.json`
2. Fill in your credentials
3. Run:
   ```bash
   npm install
   func start
   ```

## Function Details

### `ScheduledUpdateTrigger`

- **Trigger**: Timer (runs every hour at minute 5: `0 5 * * * *`)
- **Logic**:
  1. Query Supabase for enabled schedules where `day_of_week` and `time_utc` match current time
  2. For each matching schedule:
     - Authenticate as service principal
     - Get list of apps with available updates
     - Update all apps
     - Log results back to Supabase

### Schedule Matching

The function checks if a schedule should run by comparing:
- Current UTC day of week (0-6)
- Current UTC hour

A schedule marked for "Tuesday at 3:00 AM" will run when:
- `day_of_week = 2` (Tuesday)
- `time_utc = "03:00"`
- Current UTC time is between 03:00 and 03:59

## Monitoring

- View function logs in Azure Portal → Function App → **Monitor**
- Check Supabase `update_schedules` table for `last_run_at`, `last_run_status`, `last_run_result`

## Troubleshooting

### "Access Denied" when updating apps

1. Verify the service principal is added as an Application User in the Power Platform environment
2. Check the security role assigned has app management permissions
3. Ensure admin consent was granted for all API permissions

### Schedules not running

1. Check the function is running: Azure Portal → Function App → **Functions** → **ScheduledUpdateTrigger** → **Monitor**
2. Verify Supabase connection settings
3. Check that schedules have `enabled = true`

### No apps being updated

1. The function only updates apps that have available updates
2. Check if apps actually have updates pending in the UI
