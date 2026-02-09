#Requires -Version 5.1
<#
.SYNOPSIS
    Automatically configures Azure AD App Registration permissions for D365 App Updater

.DESCRIPTION
    This script adds all required API permissions to your Azure AD app registration.
    Run this once, and you're done!
    
    Prerequisites:
    - Azure CLI installed (https://aka.ms/InstallAzureCLIDocs)
    - Global Administrator or Application Administrator role

.EXAMPLE
    .\setup-azure-permissions.ps1
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$AppId,
    
    [Parameter(Mandatory=$false)]
    [string]$RedirectUri = "https://moliveirapinto.github.io/d365-app-updater/"
)

# Colors for output
$ErrorColor = "Red"
$SuccessColor = "Green"
$InfoColor = "Cyan"
$WarningColor = "Yellow"

function Write-Step {
    param([string]$Message)
    Write-Host "`n✓ $Message" -ForegroundColor $SuccessColor
}

function Write-Info {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor $InfoColor
}

function Write-Error-Custom {
    param([string]$Message)
    Write-Host "`n✗ $Message" -ForegroundColor $ErrorColor
}

function Write-Warning-Custom {
    param([string]$Message)
    Write-Host "  ⚠ $Message" -ForegroundColor $WarningColor
}

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║       D365 Power Platform App Updater Setup Script          ║
║                                                              ║
║  This will automatically configure your Azure AD app        ║
║  registration with all required API permissions.            ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor $InfoColor

# Check if Azure CLI is installed
Write-Host "Checking prerequisites..." -ForegroundColor $InfoColor
try {
    $azVersion = az version 2>&1 | Out-Null
    Write-Step "Azure CLI is installed"
} catch {
    Write-Error-Custom "Azure CLI is not installed!"
    Write-Host "`nPlease install Azure CLI from: https://aka.ms/InstallAzureCLIDocs`n" -ForegroundColor $WarningColor
    exit 1
}

# Login to Azure
Write-Host "`nChecking Azure login status..." -ForegroundColor $InfoColor
$account = az account show 2>&1 | ConvertFrom-Json
if (-not $account) {
    Write-Info "Not logged in. Opening Azure login..."
    az login --allow-no-subscriptions
    if ($LASTEXITCODE -ne 0) {
        Write-Error-Custom "Login failed. Please try again."
        exit 1
    }
    Write-Step "Successfully logged in to Azure"
} else {
    Write-Step "Already logged in as: $($account.user.name)"
}

# Get App ID if not provided
if (-not $AppId) {
    Write-Host "`n" -NoNewline
    $AppId = Read-Host "Enter your App Registration Client ID (Application ID)"
    if (-not $AppId) {
        Write-Error-Custom "App ID is required!"
        exit 1
    }
}

Write-Host "`nValidating App Registration..." -ForegroundColor $InfoColor
$app = az ad app show --id $AppId 2>&1 | ConvertFrom-Json
if (-not $app) {
    Write-Error-Custom "Could not find app with ID: $AppId"
    Write-Host "  Make sure the App ID is correct and you have access to it.`n" -ForegroundColor $WarningColor
    exit 1
}
Write-Step "Found app: $($app.displayName)"

# API Resource IDs (these are Microsoft's standard IDs)
$apiResources = @{
    "Power Platform API" = "8578e004-a5c6-46e7-913e-12f58912df43"
    "Dynamics CRM" = "00000007-0000-0000-c000-000000000000"
    "Microsoft Graph" = "00000003-0000-0000-c000-000000000000"
    "BAP/Power Platform Environment Service" = "475226c6-020e-4fb2-8a90-7a972cbfc1d4"
    "PowerApps Service" = "8109c6e9-c0b9-4c9a-a8b4-0e2a11d7a2a9"
}

# Define all required permissions with their scope IDs
$permissions = @(
    @{API="Power Platform API"; ResourceId=$apiResources["Power Platform API"]; Permission="user_impersonation"; ScopeId="9f7795e2-ce4f-42d1-b8c2-bc6c8af81d8f"}
    @{API="Dynamics CRM"; ResourceId=$apiResources["Dynamics CRM"]; Permission="user_impersonation"; ScopeId="78ce3f0f-a1ce-49c2-8cde-64b5c0896db4"}
    @{API="Microsoft Graph"; ResourceId=$apiResources["Microsoft Graph"]; Permission="User.Read"; ScopeId="e1fe6dd8-ba31-4d61-89e7-88639da4683d"}
    @{API="BAP/Power Platform Environment Service"; ResourceId=$apiResources["BAP/Power Platform Environment Service"]; Permission="user_impersonation"; ScopeId="47bdd03f-c0f9-4f02-8e03-8f58c47c2c3d"}
    @{API="PowerApps Service"; ResourceId=$apiResources["PowerApps Service"]; Permission="User"; ScopeId="2f67e4e1-3e7f-4b3e-a5ed-1e6cb0a0b4c1"}
)

Write-Host "`nAdding API Permissions..." -ForegroundColor $InfoColor

$addedCount = 0
$skippedCount = 0

foreach ($perm in $permissions) {
    Write-Info "Adding $($perm.API) - $($perm.Permission)..."
    
    $result = az ad app permission add `
        --id $AppId `
        --api $perm.ResourceId `
        --api-permissions "$($perm.ScopeId)=Scope" 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "    ✓ Added" -ForegroundColor $SuccessColor
        $addedCount++
        Start-Sleep -Milliseconds 500
    } else {
        if ($result -match "already exists") {
            Write-Host "    ⊙ Already exists" -ForegroundColor Gray
            $skippedCount++
        } else {
            Write-Warning-Custom "Could not add: $result"
        }
    }
}

Write-Step "Added $addedCount new permissions, $skippedCount already existed"

# Configure redirect URI
Write-Host "`nConfiguring Redirect URI..." -ForegroundColor $InfoColor
Write-Info "Adding: $RedirectUri"

$redirectUris = @($app.spa.redirectUris)
if ($redirectUris -notcontains $RedirectUri) {
    $redirectUris += $RedirectUri
    $redirectUrisJson = ($redirectUris | ConvertTo-Json -Compress)
    
    az ad app update --id $AppId --web-redirect-uris $redirectUrisJson 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Step "Redirect URI configured"
    } else {
        Write-Warning-Custom "Could not add redirect URI automatically. Please add manually: $RedirectUri"
    }
} else {
    Write-Host "    ⊙ Already configured" -ForegroundColor Gray
}

# Grant admin consent
Write-Host "`n" -NoNewline
$consent = Read-Host "Grant admin consent now? (requires Global Admin role) [Y/n]"
if ($consent -ne 'n' -and $consent -ne 'N') {
    Write-Info "Granting admin consent..."
    
    foreach ($perm in $permissions) {
        $result = az ad app permission admin-consent --id $AppId 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Step "Admin consent granted!"
            break
        }
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Warning-Custom "Could not grant admin consent automatically."
        Write-Host "`n  Please grant admin consent manually:" -ForegroundColor $WarningColor
        Write-Host "  1. Go to https://portal.azure.com" -ForegroundColor $WarningColor
        Write-Host "  2. Navigate to Azure AD > App Registrations > $($app.displayName)" -ForegroundColor $WarningColor
        Write-Host "  3. Go to API Permissions" -ForegroundColor $WarningColor
        Write-Host "  4. Click 'Grant admin consent for [Your Org]'" -ForegroundColor $WarningColor
    }
} else {
    Write-Warning-Custom "Skipped admin consent. You'll need to grant it manually later."
}

Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║                    ✓ SETUP COMPLETE!                        ║
║                                                              ║
║  Your app registration is now configured.                   ║
║                                                              ║
║  Next steps:                                                 ║
║  1. Go to: https://moliveirapinto.github.io/d365-app-updater/║
║  2. Enter your App ID: $AppId    
║  3. Sign in and start updating apps!                         ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝

"@ -ForegroundColor $SuccessColor

Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
