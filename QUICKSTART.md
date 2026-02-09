# Quick Start Guide

Get up and running with D365 Power Platform App Updater in 5 minutes!

## 📋 Prerequisites

- Microsoft 365 account with D365/Power Platform access
- Admin rights to create Azure AD app registrations (Global Admin or Application Admin)
- Azure CLI installed ([Download here](https://aka.ms/InstallAzureCLIDocs))
- Modern web browser (Chrome, Edge, Firefox)

## 🚀 Automated Setup (EASIEST - 2 minutes!)

### ⭐ NEW: One-Click Setup Script

**This is the easiest way!** Run our automated setup script to configure everything:

1. **Create the App Registration First:**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Click **"New registration"**
   - Name: `D365 App Updater`
   - Account type: `Single tenant`
   - Click **"Register"**
   - Copy your **Client ID** (you'll need it in step 2)

2. **Run the Setup Script:**
   - Download or clone this repository
   - Right-click `setup-azure-permissions.ps1` → **Run with PowerShell**
   - Or run manually:
   ```powershell
   cd "path\to\Update all apps"
   .\setup-azure-permissions.ps1
   ```
   - Enter your Client ID when prompted
   - The script will automatically:
     - ✅ Add all required API permissions
     - ✅ Configure redirect URIs
     - ✅ Grant admin consent
     - ✅ Verify everything is set up correctly

3. **Done!** Go to https://moliveirapinto.github.io/d365-app-updater/ and sign in!

---

## 🛠️ Manual Setup (if automated script doesn't work)

### 1️⃣ Create Azure AD App (5 minutes)

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click **"New registration"**
3. Fill in:
   - Name: `D365 App Updater`
   - Account type: `Single tenant`
4. Click **"Register"**
5. In **Authentication**:
   - Click "Add a platform" → "Single-page application"
   - Add redirect URI: `https://moliveirapinto.github.io/d365-app-updater/`
   - Enable "Access tokens" and "ID tokens"
6. In **API permissions** - Add ALL of these:
   - **Power Platform API** → user_impersonation
   - **Dynamics CRM** → user_impersonation
   - **Microsoft Graph** → User.Read
   - **BAP API** → user_impersonation
   - Click **"Grant admin consent for [Your Org]"** (IMPORTANT!)
7. Copy your **Client ID** and **Tenant ID** from the Overview page

📖 **Need detailed instructions?** See [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md)

### 2️⃣ Use the App (1 minute)

#### Option A: Using PowerShell (Recommended for Windows)

```powershell
cd "path\to\Update all apps"
.\start-dev-server.ps1
```

#### Option B: Using Python

```powershell
cd "path\to\Update all apps"
python -m http.server 8000
```

#### Option C: Using Node.js

```powershell
cd "path\to\Update all apps"
npm install
npm start
```

### 3️⃣ Connect and Use (1 minute)

1. Open http://localhost:8000 in your browser
2. Enter your credentials:
   - **Organization URL**: `https://yourorg.crm.dynamics.com`
   - **Tenant ID**: [from step 1]
   - **Client ID**: [from step 1]
3. Click **"Connect to Power Platform"**
4. Sign in when prompted
5. View and update your apps! 🎉

## 🎯 What You Can Do

✅ **View all installed apps** in your environment  
✅ **See which apps have updates** available  
✅ **Update individual apps** one at a time  
✅ **Update all apps at once** (the main feature!)  
✅ **Save credentials** for quick access  

## 🔧 Troubleshooting

### Can't authenticate?
- ✓ Check that redirect URI matches exactly: `http://localhost:8000`
- ✓ Verify app is configured as "Single-page application" (not Web)
- ✓ Ensure admin consent is granted for Dynamics CRM permission

### Can't see apps?
- ✓ Verify your organization URL is correct
- ✓ Check that you have admin access to the environment
- ✓ Try refreshing the apps list

### Server won't start?
- If using Python: Install from https://www.python.org/downloads/
- If using Node.js: Run `npm install` first
- If using PowerShell: Right-click the .ps1 file → "Run with PowerShell"

## 📚 Next Steps

- [ ] Read the full [README.md](README.md) for detailed information
- [ ] Check [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md) for detailed Azure configuration
- [ ] Review [POWERPLATFORM_API.md](POWERPLATFORM_API.md) for implementing real updates
- [ ] Deploy to Azure Static Web Apps or GitHub Pages for production use

## 💡 Tips

- **Check the "Remember me" box** to save credentials between sessions
- **Always test in a development environment first** before production
- **Use the refresh button** to reload the app list after making changes
- **Watch the browser console** for detailed error messages if something goes wrong

## 🆘 Need Help?

1. Check the browser console (F12) for error messages
2. Review the troubleshooting section above
3. Read the full documentation in this repository
4. Check Azure AD sign-in logs for authentication issues

## ⚠️ Important Notes

- **Current Version**: The update functionality uses simulated data
- **Real Updates**: See [POWERPLATFORM_API.md](POWERPLATFORM_API.md) for implementing actual updates
- **Security**: Always test in development before production use
- **Backup**: Ensure you have backups before updating critical apps

---

**Estimated Time**: 5 minutes  
**Difficulty**: Easy  
**Prerequisites**: Azure AD access, D365 environment

Happy updating! 🚀
