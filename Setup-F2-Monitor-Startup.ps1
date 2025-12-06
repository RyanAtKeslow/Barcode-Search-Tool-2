# Setup script to run F2 Desktop Monitor at Windows startup
# Run this script once to set up automatic startup

$scriptPath = Join-Path $PSScriptRoot "F2-Desktop-Monitor.ps1"
$taskName = "F2 Desktop Monitor"

Write-Host "Setting up F2 Desktop Monitor to run at startup..." -ForegroundColor Cyan
Write-Host "Script path: $scriptPath" -ForegroundColor Gray

# Check if script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "Error: Script not found at $scriptPath" -ForegroundColor Red
    exit 1
}

# Remove existing task if it exists
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create scheduled task
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Automatically monitors Desktop for F2 export files and moves them to Google Drive"

Write-Host ""
Write-Host "Success! F2 Desktop Monitor is now set to run at startup." -ForegroundColor Green
Write-Host "The script will run in the background (no window visible)." -ForegroundColor Green
Write-Host ""
Write-Host "To test it now, run:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName `"$taskName`"" -ForegroundColor Gray
Write-Host ""
Write-Host "To remove the startup task later, run:" -ForegroundColor Yellow
Write-Host "  Unregister-ScheduledTask -TaskName `"$taskName`" -Confirm:`$false" -ForegroundColor Gray

