# Simple setup script - Uses Windows Startup folder (no admin required!)
# This is the easiest and most reliable method for user-level startup

$scriptPath = Join-Path $PSScriptRoot "F2-Desktop-Monitor.ps1"
$startupFolder = [System.Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "F2 Desktop Monitor.lnk"

Write-Host "Setting up F2 Desktop Monitor to run at startup..." -ForegroundColor Cyan
Write-Host "Script path: $scriptPath" -ForegroundColor Gray
Write-Host "Startup folder: $startupFolder" -ForegroundColor Gray
Write-Host ""

# Check if script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "❌ Error: Script not found at $scriptPath" -ForegroundColor Red
    exit 1
}

# Remove existing shortcut if it exists
if (Test-Path $shortcutPath) {
    Write-Host "Removing existing shortcut..." -ForegroundColor Yellow
    Remove-Item $shortcutPath -Force
}

# Create shortcut using WScript.Shell COM object
try {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.Description = "F2 Desktop Monitor - Automatically moves F2 export files to Google Drive"
    $shortcut.Save()
    
    Write-Host "✅ Success! F2 Desktop Monitor is now set to run at startup." -ForegroundColor Green
    Write-Host "The script will run in the background (no window visible)." -ForegroundColor Green
    Write-Host ""
    Write-Host "Shortcut created at: $shortcutPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "To test it now, you can:" -ForegroundColor Yellow
    Write-Host "  1. Run the script directly: .\F2-Desktop-Monitor.ps1" -ForegroundColor Cyan
    Write-Host "  2. Or double-click the shortcut in the Startup folder" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To remove the startup entry later:" -ForegroundColor Yellow
    Write-Host "  Remove-Item `"$shortcutPath`"" -ForegroundColor Gray
    
} catch {
    Write-Host "❌ Error creating shortcut: $_" -ForegroundColor Red
    exit 1
}

