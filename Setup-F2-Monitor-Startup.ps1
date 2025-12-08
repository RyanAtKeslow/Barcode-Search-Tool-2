# Setup script to run F2 Desktop Monitor at Windows startup
# Run this script once to set up automatic startup
# NOTE: This script creates a USER-LEVEL task that runs for YOUR account
# You do NOT need Administrator privileges - just run it normally!

$scriptPath = Join-Path $PSScriptRoot "F2-Desktop-Monitor.ps1"
$taskName = "F2 Desktop Monitor"

Write-Host "Setting up F2 Desktop Monitor to run at startup..." -ForegroundColor Cyan
Write-Host "Script path: $scriptPath" -ForegroundColor Gray
Write-Host "User: $env:USERNAME" -ForegroundColor Gray
Write-Host ""

# Check if script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "Error: Script not found at $scriptPath" -ForegroundColor Red
    Write-Host "Looking for: F2-Desktop-Monitor.ps1" -ForegroundColor Yellow
    exit 1
}

# Get current user's SID for proper user-level task creation
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$userSid = $currentUser.User.Value
$userName = $env:USERNAME

Write-Host "Creating task for user: $userName" -ForegroundColor Cyan
Write-Host ""

# Remove existing task if it exists (check both root and user folder)
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
    } catch {
        # Try with full path
        $taskPath = $existingTask.TaskPath
        Unregister-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    }
}

# Create scheduled task - USER LEVEL (no admin required)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
# Use current user with Interactive logon (runs when user logs in)
$principal = New-ScheduledTaskPrincipal -UserId $userName -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

# Use root task path (works for user-level tasks)
$taskPath = "\"

try {
    # Try to create task in root path (user-level, no admin required)
    Register-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Automatically monitors Desktop for F2 export files and moves them to Google Drive" -ErrorAction Stop
    
    # Verify task was created
    $createdTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($createdTask) {
        Write-Host "✅ Task created successfully!" -ForegroundColor Green
        Write-Host "Task State: $($createdTask.State)" -ForegroundColor Gray
        Write-Host "Task Path: $($createdTask.TaskPath)" -ForegroundColor Gray
        Write-Host "Principal: $($createdTask.Principal.UserId)" -ForegroundColor Gray
    } else {
        Write-Host "⚠️ Warning: Task may not have been created properly" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Error creating scheduled task: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Trying alternative method using schtasks.exe..." -ForegroundColor Yellow
    
    # Alternative: Use schtasks.exe which works well for user-level tasks
    # schtasks requires the entire /TR value to be in quotes when it contains spaces
    # We need to escape quotes properly: use "" for literal quote inside quoted string
    $taskCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""$scriptPath"""
    
    # Create user-level task using schtasks (no admin required)
    # Use /RU to specify current user explicitly
    $schtasksArgs = @(
        "/CREATE",
        "/TN", $taskName,
        "/TR", $taskCommand,
        "/SC", "ONLOGON",
        "/RL", "LIMITED",
        "/RU", $userName,
        "/F"
    )
    
    Write-Host "Running: schtasks /CREATE /TN `"$taskName`" /TR `"$taskCommand`" /SC ONLOGON /RL LIMITED /RU $userName /F" -ForegroundColor Gray
    $result = & schtasks.exe $schtasksArgs 2>&1
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -eq 0) {
        Write-Host "✅ Task created successfully using schtasks.exe!" -ForegroundColor Green
        Start-Sleep -Seconds 1
        $createdTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($createdTask) {
            Write-Host "Task State: $($createdTask.State)" -ForegroundColor Gray
            Write-Host "Task Path: $($createdTask.TaskPath)" -ForegroundColor Gray
            Write-Host "Principal: $($createdTask.Principal.UserId)" -ForegroundColor Gray
        }
    } else {
        Write-Host "❌ Failed to create task using schtasks.exe" -ForegroundColor Red
        Write-Host "Output: $result" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting:" -ForegroundColor Yellow
        Write-Host "  1. Make sure PowerShell execution policy allows scripts:" -ForegroundColor Cyan
        Write-Host "     Get-ExecutionPolicy (should be RemoteSigned or Unrestricted)" -ForegroundColor Gray
        Write-Host "  2. Try running: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Cyan
        Write-Host "  3. Verify the script path is correct: $scriptPath" -ForegroundColor Cyan
        exit 1
    }
}

Write-Host ""
Write-Host "Success! F2 Desktop Monitor is now set to run at startup." -ForegroundColor Green
Write-Host "The script will run in the background (no window visible)." -ForegroundColor Green
Write-Host ""
Write-Host "To test it now, run:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName `"$taskName`"" -ForegroundColor Gray
Write-Host ""
Write-Host "To remove the startup task later, run:" -ForegroundColor Yellow
Write-Host "  Unregister-ScheduledTask -TaskName `"$taskName`" -Confirm:`$false" -ForegroundColor Gray

