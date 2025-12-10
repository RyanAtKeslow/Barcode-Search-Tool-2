# F2 Desktop Monitor - Moves F2 files to Google Drive sync folder
# This script works with Google Drive for Desktop.
# It moves files from Desktop to a local Google Drive folder,
# and Google Drive for Desktop handles the sync automatically.

# Configuration
$DESKTOP_PATH = "$env:USERPROFILE\Desktop"
$SYNC_FOLDER_PATH = "G:\My Drive\Service Board Imports"
$FILE_PATTERN = "^Service \d{4}-\d{2}-\d{2} at \d{1,2}\.\d{2}(\.\d{2})? (AM|PM)\.xlsx$"
$CHECK_INTERVAL = 10
$PROCESSED_FILES_LOG = "$env:APPDATA\F2DesktopMonitor\processed_files.txt"
$LOG_FILE = "$env:APPDATA\F2DesktopMonitor\monitor.log"

# Google Apps Script Web App URL for triggering F2 import processing
# To get this URL:
# 1. Open the Apps Script project in Google Apps Script editor
# 2. Click "Deploy" > "New deployment" > "Web app"
# 3. Execute as: "Me", Who has access: "Anyone" (or "Anyone with Google account")
# 4. Click "Deploy" and copy the Web app URL here
# 5. Leave empty to disable automatic processing trigger
$F2_IMPORT_WEB_APP_URL = "https://script.google.com/a/macros/keslowcamera.com/s/AKfycbzJHTZn1qVexotPU2klRroIPWNKYxUn-VfZ34Xcf72fyblp8dGa4gvD-fgUmZy5q8or/exec"  # TODO: Add your web app URL here

# Create log directory if it doesn't exist
$logDir = Split-Path -Parent $PROCESSED_FILES_LOG
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LOG_FILE -Value $logMessage -ErrorAction SilentlyContinue
    # Also write to console if available (for manual runs)
    if ($Host.UI.RawUI) {
        switch ($Level) {
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "WARN" { Write-Host $logMessage -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
            default { Write-Host $logMessage -ForegroundColor Cyan }
        }
    }
}

# Check if sync folder exists
if (-not (Test-Path $SYNC_FOLDER_PATH)) {
    Write-Log "Sync folder not found: $SYNC_FOLDER_PATH" "ERROR"
    Write-Log "Please update SYNC_FOLDER_PATH in the script to your Google Drive sync folder" "ERROR"
    exit 1
}

# Load processed files list
function Get-ProcessedFiles {
    if (Test-Path $PROCESSED_FILES_LOG) {
        return @(Get-Content $PROCESSED_FILES_LOG | Where-Object { $_ -ne "" })
    }
    return @()
}

# Add file to processed list
function Add-ProcessedFile {
    param([string]$fileName)
    $processed = Get-ProcessedFiles
    if ($processed -notcontains $fileName) {
        Add-Content -Path $PROCESSED_FILES_LOG -Value $fileName
    }
}

# Check if file matches F2 pattern
function Test-F2FilePattern {
    param([string]$fileName)
    return $fileName -match $FILE_PATTERN
}

# Trigger F2 import processing via Google Apps Script web app
function Invoke-F2Import {
    param([string]$fileName, [string]$webAppUrl)
    
    # Check if web app URL is configured
    if ([string]::IsNullOrWhiteSpace($webAppUrl)) {
        Write-Log "F2 Import web app URL not configured, skipping automatic processing trigger" "WARN"
        return $false
    }
    
    try {
        Write-Log "Waiting 30 seconds before triggering F2 import processing for $fileName..." "INFO"
        Start-Sleep -Seconds 30
        
        Write-Log "Triggering F2 import processing via web app..." "INFO"
        $response = Invoke-WebRequest -Uri $webAppUrl -Method Get -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
        
        if ($response.StatusCode -eq 200) {
            try {
                $result = $response.Content | ConvertFrom-Json
                if ($result.success) {
                    Write-Log "F2 import processing triggered successfully" "SUCCESS"
                    return $true
                } else {
                    Write-Log "F2 import processing returned error: $($result.error)" "ERROR"
                    return $false
                }
            }
            catch {
                # Response might not be JSON, but status 200 means it worked
                Write-Log "F2 import processing triggered (response: $($response.Content))" "SUCCESS"
                return $true
            }
        } else {
            Write-Log "F2 import web app returned status code: $($response.StatusCode)" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error triggering F2 import processing: $_" "ERROR"
        return $false
    }
}

# Send Windows toast notification
function Send-Notification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$AppId = "F2DesktopMonitor"
    )
    
    try {
        # Try using BurntToast module if available
        if (Get-Module -ListAvailable -Name BurntToast) {
            Import-Module BurntToast -ErrorAction SilentlyContinue
            New-BurntToastNotification -Text $Title, $Message -AppId $AppId -ErrorAction Stop
            Write-Log "Notification sent via BurntToast" "INFO"
            return
        }
        
        # Fallback: Use Windows.UI.Notifications API (Windows 10/11)
        try {
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
            
            # Use PowerShell's AppId for better compatibility, or use a proper AUMID format
            # Try using PowerShell's actual AppId first
            $powerShellAppId = "Microsoft.Windows.Shell.RunDialog"
            $usePowerShellAppId = $false
            
            # Check if we can use PowerShell's AppId
            try {
                [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($powerShellAppId) | Out-Null
                $usePowerShellAppId = $true
            }
            catch {
                # Fall back to custom AppId
                $usePowerShellAppId = $false
            }
            
            $effectiveAppId = if ($usePowerShellAppId) { $powerShellAppId } else { $AppId }
            
            # Enhanced toast template with audio and longer duration
            $template = @"
<toast scenario="reminder" duration="long" useButtonStyle="true">
    <visual>
        <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default" />
</toast>
"@
            
            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($template)
            
            $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
            
            # Set expiration time to 10 seconds (default is 7 seconds)
            $toast.ExpirationTime = [DateTimeOffset]::Now.AddSeconds(10)
            
            $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($effectiveAppId)
            $notifier.Show($toast)
            Write-Log "Notification sent via Windows.UI.Notifications (AppId: $effectiveAppId)" "INFO"
            return
        }
        catch {
            Write-Log "Windows.UI.Notifications failed: $_" "WARN"
            throw  # Re-throw to trigger MessageBox fallback
        }
    }
    catch {
        # If all else fails, use a simple popup (requires user interaction to dismiss)
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            Write-Log "Notification sent via MessageBox fallback" "INFO"
        }
        catch {
            # Last resort: just log it
            Write-Log "All notification methods failed: $_" "ERROR"
        }
    }
}

# Check if a file is locked/in use by another process
function Test-FileLocked {
    param([string]$filePath)
    
    try {
        # Try to open the file for exclusive access
        $fileStream = [System.IO.File]::Open($filePath, 'Open', 'ReadWrite', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        return $false
    }
    catch {
        # File is locked if we can't open it
        return $true
    }
}

# Check if Excel has the file open
function Test-ExcelHasFileOpen {
    param([string]$filePath)
    
    try {
        $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
        if ($excelProcesses) {
            # Check if any Excel process has this file open
            # This is a best-effort check - we can't directly query what files Excel has open
            # But if Excel is running and file is locked, it's likely Excel has it
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

# Move file with retry logic for file locks
function Move-FileWithRetry {
    param(
        [string]$filePath,
        [string]$destPath,
        [string]$fileName,
        [int]$maxRetries = 6,
        [int]$initialWaitSeconds = 5
    )
    
    $attempt = 0
    $waitSeconds = $initialWaitSeconds
    
    while ($attempt -lt $maxRetries) {
        $attempt++
        
        # Check if file is locked
        if (Test-FileLocked -filePath $filePath) {
            $excelOpen = Test-ExcelHasFileOpen -filePath $filePath
            if ($excelOpen) {
                Write-Log "File $fileName is locked (likely by Excel). Waiting $waitSeconds seconds before retry $attempt/$maxRetries..." "WARN"
            } else {
                Write-Log "File $fileName is locked by another process. Waiting $waitSeconds seconds before retry $attempt/$maxRetries..." "WARN"
            }
            
            if ($attempt -lt $maxRetries) {
                Start-Sleep -Seconds $waitSeconds
                # Exponential backoff: increase wait time with each retry
                $waitSeconds = [Math]::Min($waitSeconds * 1.5, 30)
            }
            continue
        }
        
        # Try to move the file
        try {
            Move-Item -Path $filePath -Destination $destPath -Force -ErrorAction Stop
            Write-Log "Successfully moved $fileName" "SUCCESS"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            if ($errorMsg -like "*being used by another process*" -or $errorMsg -like "*cannot access*") {
                Write-Log "File $fileName is in use. Waiting $waitSeconds seconds before retry $attempt/$maxRetries..." "WARN"
                if ($attempt -lt $maxRetries) {
                    Start-Sleep -Seconds $waitSeconds
                    $waitSeconds = [Math]::Min($waitSeconds * 1.5, 30)
                }
                continue
            } else {
                # Different error - don't retry
                Write-Log "Error moving file $fileName : $errorMsg" "ERROR"
                return $false
            }
        }
    }
    
    # All retries exhausted
    Write-Log "Failed to move $fileName after $maxRetries attempts. File may still be open in Excel or another application." "ERROR"
    return $false
}

# Main monitoring loop
Write-Log "F2 Desktop Monitor started" "INFO"
Write-Log "Monitoring: $DESKTOP_PATH" "INFO"
Write-Log "Sync folder: $SYNC_FOLDER_PATH" "INFO"
Write-Log "Check interval: $CHECK_INTERVAL seconds" "INFO"

$processedFiles = Get-ProcessedFiles

while ($true) {
    try {
        # Get all Excel files on desktop
        $excelFiles = Get-ChildItem -Path $DESKTOP_PATH -Filter "*.xlsx" -ErrorAction SilentlyContinue
        
        foreach ($file in $excelFiles) {
            $fileName = $file.Name
            $filePath = $file.FullName
            $destPath = Join-Path $SYNC_FOLDER_PATH $fileName
            
            # Skip if already processed
            if ($processedFiles -contains $fileName) {
                continue
            }
            
            # Check if file matches F2 pattern
            if (-not (Test-F2FilePattern $fileName)) {
                continue
            }
            
            # Check if file is still being written (recently modified)
            # Increased wait time since Excel may auto-open the file after export
            $timeSinceModified = (Get-Date) - $file.LastWriteTime
            if ($timeSinceModified.TotalSeconds -lt 15) {
                Write-Log "File $fileName was recently created/modified (${([int]$timeSinceModified.TotalSeconds)}s ago), waiting for Excel to finish opening..." "WARN"
                continue
            }
            
            # Check if file already exists in destination
            if (Test-Path $destPath) {
                Write-Log "File $fileName already exists in sync folder, skipping..." "WARN"
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
                continue
            }
            
            # Move file to sync folder with retry logic
            Write-Log "Moving $fileName to sync folder..." "INFO"
            if (Move-FileWithRetry -filePath $filePath -destPath $destPath -fileName $fileName) {
                # Mark as processed only if move succeeded
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
                
                # Send system notification
                Write-Log "Sending notification for $fileName..." "INFO"
                Send-Notification -Title "F2 File Moved" -Message "Successfully moved $fileName to Google Drive sync folder"
                
                # Trigger F2 import processing after 5 second delay (runs in background job)
                Start-Job -ScriptBlock {
                    param($webAppUrl, $fileName, $logFile)
                    
                    # Simple logging function for background job
                    function Write-Log {
                        param([string]$Message, [string]$Level = "INFO")
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $logMessage = "[$timestamp] [$Level] $Message"
                        Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
                    }
                    
                    # Check if web app URL is configured
                    if ([string]::IsNullOrWhiteSpace($webAppUrl)) {
                        Write-Log "F2 Import web app URL not configured, skipping automatic processing trigger" "WARN"
                        return
                    }
                    
                    try {
                        Write-Log "Waiting 5 seconds before triggering F2 import processing for $fileName..." "INFO"
                        Start-Sleep -Seconds 5
                        
                        Write-Log "Triggering F2 import processing via web app..." "INFO"
                        $response = Invoke-WebRequest -Uri $webAppUrl -Method Get -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
                        
                        if ($response.StatusCode -eq 200) {
                            try {
                                $result = $response.Content | ConvertFrom-Json
                                if ($result.success) {
                                    Write-Log "F2 import processing triggered successfully" "SUCCESS"
                                } else {
                                    Write-Log "F2 import processing returned error: $($result.error)" "ERROR"
                                }
                            }
                            catch {
                                # Response might not be JSON, but status 200 means it worked
                                Write-Log "F2 import processing triggered (response: $($response.Content))" "SUCCESS"
                            }
                        } else {
                            Write-Log "F2 import web app returned status code: $($response.StatusCode)" "ERROR"
                        }
                    }
                    catch {
                        $errorDetails = $_.Exception.Message
                        if ($_.Exception.Response) {
                            $statusCode = $_.Exception.Response.StatusCode.value__
                            $errorDetails += " (HTTP $statusCode)"
                        }
                        Write-Log "Error triggering F2 import processing: $errorDetails" "ERROR"
                        Write-Log "Web app URL used: $webAppUrl" "ERROR"
                    }
                } -ArgumentList $F2_IMPORT_WEB_APP_URL, $fileName, $LOG_FILE
                
                # Clean up completed jobs periodically (don't wait for this one)
                Get-Job | Where-Object { $_.State -eq 'Completed' -or $_.State -eq 'Failed' } | Remove-Job -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        Write-Log "Error in monitoring loop: $_" "ERROR"
    }
    
    # Wait before next check
    Start-Sleep -Seconds $CHECK_INTERVAL
}
