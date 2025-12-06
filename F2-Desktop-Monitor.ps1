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

# Create log directory if it doesn't exist
$logDir = Split-Path -Parent $PROCESSED_FILES_LOG
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Check if sync folder exists
if (-not (Test-Path $SYNC_FOLDER_PATH)) {
    Write-Host "Sync folder not found: $SYNC_FOLDER_PATH" -ForegroundColor Yellow
    Write-Host "Please update SYNC_FOLDER_PATH in the script to your Google Drive sync folder" -ForegroundColor Yellow
    exit 1
}

# Load processed files list
function Get-ProcessedFiles {
    if (Test-Path $PROCESSED_FILES_LOG) {
        return Get-Content $PROCESSED_FILES_LOG | Where-Object { $_ -ne "" }
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
Write-Host "F2 Desktop Monitor started" -ForegroundColor Green
Write-Host "Monitoring: $DESKTOP_PATH" -ForegroundColor Cyan
Write-Host "Sync folder: $SYNC_FOLDER_PATH" -ForegroundColor Cyan
Write-Host "Check interval: $CHECK_INTERVAL seconds" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Ctrl+C to stop monitoring" -ForegroundColor Yellow
Write-Host ""

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
                Write-Host "File $fileName is still being written, waiting..." -ForegroundColor Yellow
                continue
            }
            
            # Check if file already exists in destination
            if (Test-Path $destPath) {
                Write-Host "File $fileName already exists in sync folder, skipping..." -ForegroundColor Yellow
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
                continue
            }
            
            # Move file to sync folder
            try {
                Write-Host "Moving $fileName to sync folder..." -ForegroundColor Cyan
                Move-Item -Path $filePath -Destination $destPath -Force
                Write-Host "Successfully moved $fileName" -ForegroundColor Green
                
                # Mark as processed
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
            }
            catch {
                Write-Host "Error moving file: $_" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "Error in monitoring loop: $_" -ForegroundColor Red
    }
    
    # Wait before next check
    Start-Sleep -Seconds $CHECK_INTERVAL
}
