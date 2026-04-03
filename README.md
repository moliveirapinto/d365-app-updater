# D365 Power Platform App Updater

A web-based tool to manage and update Microsoft Dynamics 365 / Power Platform applications in bulk. This tool addresses the limitation in the Power Platform Admin Center where you can only update apps one at a time.

You can access the live version here: `https://moliveirapinto.github.io/d365-app-updater/`


## 🎯 Features

- ✅ **Bulk App Updates**: Update all available apps at once
- ✅ **Multi-Select Updates**: Select specific apps to update
- ✅ **Environment Switcher**: Quickly switch between environments
- ✅ **Individual Updates**: Update specific apps one by one
- ✅ **Authentication**: Secure MSAL-based authentication
- ✅ **Update Detection**: Automatically detects which apps have updates available
- ✅ **User-Friendly Interface**: Clean, modern Bootstrap UI
- ✅ **Session Management**: Persistent login across sessions
- ✅ **Admin Dashboard**: Track usage analytics with Supabase
- ✅ **Real-time Status**: Live update status tracking
- ✅ **Scheduled Auto-Updates**: Configure automatic updates on a weekly schedule (requires Azure Function)

## 🚀 Getting Started

### Prerequisites

1. **Azure AD App Registration**: You need to create an Azure AD app registration for authentication
2. **Permissions**: Admin access to your Power Platform environment
3. **Modern Browser**: Chrome, Edge, Firefox, or Safari

### Azure AD Setup

⚡ **EASIEST WAY: Use Our Automated Setup Script!**

1. Create an app registration in [Azure Portal](https://portal.azure.com) → Azure AD → App registrations
   - Name: `D365 App Updater`
   - Account type: Single tenant
   - Click Register

2. Run our automated setup script:
   ```powershell
   .\setup-azure-permissions.ps1
   ```
   - It will automatically configure ALL required API permissions
   - It will add redirect URIs
   - It will grant admin consent
   - **Done in 30 seconds!**

📖 **See [QUICKSTART.md](QUICKSTART.md) for step-by-step instructions**

---

**Manual Setup (if script doesn't work):**

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** → **App registrations**
3. Click **"New registration"**
4. Configure:
   - **Name**: D365 App Updater
   - **Supported account types**: Single tenant
   - **Redirect URI**: 
     - Platform: **Single-page application (SPA)**
     - URI: `https://moliveirapinto.github.io/d365-app-updater/`

5. After registration:
   - Go to **Authentication** → Enable "Access tokens" and "ID tokens"
   - Go to **API permissions** → Add ALL these permissions:
     - **Power Platform API** → user_impersonation
     - **Dynamics CRM** → user_impersonation  
     - **Microsoft Graph** → User.Read
     - **BAP API** → user_impersonation
   - Click **"Grant admin consent for [Your Org]"** ⚠️ Required!

6. Copy your **Application (client) ID**

📖 **Detailed manual setup:** [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md)

## 📦 Installation

### Option 1: Run Locally

1. Clone or download this repository
2. Open a terminal in the project folder
3. Start a local web server:

```powershell
# Using Python (Python 3)
python -m http.server 8000

# Or using Node.js (if you have http-server installed)
npx http-server -p 8000
```

4. Open your browser and navigate to `http://localhost:8000`

### Option 2: Deploy to Azure Static Web Apps

1. Create an Azure Static Web App
2. Upload the files (`index.html`, `app.js`)
3. Configure the redirect URI in your Azure AD app to match your Static Web App URL

### Option 3: Deploy to GitHub Pages

1. Create a GitHub repository
2. Push the files to the repository
3. Enable GitHub Pages in repository settings
4. Update the redirect URI in your Azure AD app

## 🔧 Usage

1. **Open the Application**: Navigate to your deployed URL or local server
2. **Enter Credentials**:
   - Organization URL: `https://yourorg.crm.dynamics.com`
   - Client ID: Your App Registration Client ID (GUID)
3. **Connect**: Click "Connect to Power Platform"
4. **Authenticate**: Sign in with your Microsoft account in the popup
5. **View Apps**: See all installed applications and their update status
6. **Update**: 
   - Click "Update Now" for individual apps
   - Click "Update All" to update all apps at once

## 📁 File Structure

```
.
├── index.html          # Main HTML page
├── app.js             # JavaScript application logic
├── azure-function/    # Azure Function for scheduled updates
└── README.md          # This file
```

## ⏰ Scheduled Auto-Updates (Optional)

You can configure automatic weekly updates that run without user intervention. This requires:

1. **Supabase** - Store schedule configurations (see [SUPABASE_SETUP.md](SUPABASE_SETUP.md))
2. **Azure Function** - Backend service that executes updates on schedule

### How it works

1. In the app, toggle on "Auto-Update Schedule" and select your preferred day/time
2. The schedule is saved to Supabase
3. An Azure Function checks for due schedules every hour
4. When a schedule is due, the function authenticates and updates all apps

### Setup

See [azure-function/README.md](azure-function/README.md) for detailed setup instructions.

## 🔒 Security Notes

- All authentication is handled via Microsoft's MSAL library
- No credentials are stored on any server
- Access tokens are stored in browser session storage
- Credentials can optionally be saved in browser local storage (encrypted by browser)

## ⚙️ Technical Details

### Technologies Used

- **MSAL.js 2.x**: Microsoft Authentication Library for browser-based authentication
- **Bootstrap 5**: UI framework
- **Font Awesome 6**: Icons
- **Dynamics 365 Web API**: For querying and updating applications

### API Endpoints Used

The application uses the following Dynamics 365 Web API endpoints:

- `WhoAmI`: Test connection and get user info
- `organizations`: Get environment information
- `msdyn_solutions` or `solutions`: Query installed applications
- Additional endpoints for triggering updates (to be implemented)

## 🐛 Troubleshooting

### Authentication Errors

**Error: "AADSTS9002326"**
- **Cause**: App is not configured as SPA
- **Solution**: In Azure AD, remove Web platform and add SPA platform

**Error: "AADSTS500113"**
- **Cause**: Redirect URI not configured
- **Solution**: Add your app URL to the redirect URIs in Azure AD

**Error: "AADSTS65001"**
- **Cause**: No admin consent granted
- **Solution**: Grant admin consent for the Dynamics CRM API permission

### Connection Issues

**Cannot connect to organization**
- Verify your organization URL is correct
- Ensure you have access to the environment
- Check that API permissions are granted

## 🚧 Current Limitations

1. **Update Simulation**: The actual update mechanism is currently simulated. Real implementation requires:
   - Power Platform API calls for update detection
   - Installation triggering via admin APIs
   - Status polling for completion

2. **Update Detection**: Currently uses mock data. Real implementation needs:
   - Query Power Platform catalog for available versions
   - Compare with installed versions
   - Fetch update details

## 🔮 Future Enhancements

- [ ] Real Power Platform API integration for updates
- [ ] Progress bars for individual app updates
- [ ] Update history and logs
- [ ] Scheduled updates
- [ ] Email notifications
- [ ] Export/import app lists
- [ ] Dark mode support
- [ ] Multi-environment support

## 📝 Notes

This tool was created based on the authentication pattern from the [d365-datagen](https://github.com/moliveirapinto/d365-datagen) repository, which provides excellent MSAL authentication examples for D365 environments.

## 🤝 Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## 📄 License

This project is provided as-is for educational and productivity purposes.

## ⚠️ Disclaimer

This tool interacts with your Power Platform environment. Always test in a development environment first. The authors are not responsible for any issues that may arise from using this tool.

## 📞 Support

For issues related to:
- **Azure AD Setup**: Refer to [Microsoft Documentation](https://docs.microsoft.com/azure/active-directory/)
- **Power Platform**: Check [Power Platform Documentation](https://docs.microsoft.com/power-platform/)
- **This Tool**: Open an issue in the repository

---

Made with ❤️ for the Power Platform community
