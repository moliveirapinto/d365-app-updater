# D365 Power Platform App Updater - Local Development Server
# This script starts a simple HTTP server for local testing

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " D365 App Updater - Development Server" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
$pythonInstalled = Get-Command python -ErrorAction SilentlyContinue

if ($pythonInstalled) {
    Write-Host "✓ Python detected" -ForegroundColor Green
    Write-Host ""
    Write-Host "Starting server on http://localhost:8000" -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Important:" -ForegroundColor Red
    Write-Host "  Make sure your Azure AD app has this redirect URI:" -ForegroundColor White
    Write-Host "  http://localhost:8000" -ForegroundColor Cyan
    Write-Host ""
    
    # Start Python HTTP server
    python -m http.server 8000
}
else {
    Write-Host "✗ Python not found" -ForegroundColor Red
    Write-Host ""
    Write-Host "Alternative methods to run the server:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Using Node.js (if installed):" -ForegroundColor White
    Write-Host "   npx http-server -p 8000" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "2. Using VS Code:" -ForegroundColor White
    Write-Host "   - Install 'Live Server' extension" -ForegroundColor Cyan
    Write-Host "   - Right-click index.html and select 'Open with Live Server'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "3. Install Python:" -ForegroundColor White
    Write-Host "   - Download from: https://www.python.org/downloads/" -ForegroundColor Cyan
    Write-Host "   - Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Cyan
    Write-Host ""
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
