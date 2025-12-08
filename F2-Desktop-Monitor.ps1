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
            $timeSinceModified = (Get-Date) - $file.LastWriteTime
            if ($timeSinceModified.TotalSeconds -lt 5) {
                Write-Log "File $fileName is still being written, waiting..." "WARN"
                continue
            }
            
            # Check if file already exists in destination
            if (Test-Path $destPath) {
                Write-Log "File $fileName already exists in sync folder, skipping..." "WARN"
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
                continue
            }
            
            # Move file to sync folder
            try {
                Write-Log "Moving $fileName to sync folder..." "INFO"
                Move-Item -Path $filePath -Destination $destPath -Force
                Write-Log "Successfully moved $fileName" "SUCCESS"
                
                # Mark as processed
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
            }
            catch {
                Write-Log "Error moving file: $_" "ERROR"
            }
        }
    }
    catch {
        Write-Log "Error in monitoring loop: $_" "ERROR"
    }
    
    # Wait before next check
    Start-Sleep -Seconds $CHECK_INTERVAL
}
