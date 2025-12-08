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
            }
        }
    }
    catch {
        Write-Log "Error in monitoring loop: $_" "ERROR"
    }
    
    # Wait before next check
    Start-Sleep -Seconds $CHECK_INTERVAL
}
