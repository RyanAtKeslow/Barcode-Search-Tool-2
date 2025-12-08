# F2 Desktop Monitor - Automatically uploads F2 export files from Desktop to Google Drive
# 
# This script monitors your Desktop for new F2 Service Board export files and
# automatically uploads them to your Google Drive "Service Board Imports" folder.
#
# Setup Instructions:
# 1. Install Google Drive API PowerShell module: Install-Module -Name GoogleDriveAPI
# 2. Set up OAuth credentials (see README_F2_MONITOR.md)
# 3. Update the GOOGLE_DRIVE_FOLDER_ID below
# 4. Run this script (or set it to run at startup)

# Configuration
$GOOGLE_DRIVE_FOLDER_ID = "1nUy7lWNr1BVCAxyLsnFASCTszQpjgEnd"  # Service Board Imports folder
$DESKTOP_PATH = "$env:USERPROFILE\Desktop"
$FILE_PATTERN = "^Service \d{4}-\d{2}-\d{2} at \d{1,2}\.\d{2}(\.\d{2})? (AM|PM)\.xlsx$"
$CHECK_INTERVAL = 30  # Check every 30 seconds
$PROCESSED_FILES_LOG = "$env:APPDATA\F2DesktopMonitor\processed_files.txt"

# Create log directory if it doesn't exist
$logDir = Split-Path -Parent $PROCESSED_FILES_LOG
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
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

# Upload file to Google Drive using rclone (simpler than OAuth)
function Upload-ToGoogleDrive {
    param(
        [string]$filePath,
        [string]$fileName
    )
    
    try {
        # Check if rclone is installed
        $rclonePath = Get-Command rclone -ErrorAction SilentlyContinue
        if (-not $rclonePath) {
            Write-Host "❌ rclone is not installed. Please install it from https://rclone.org/" -ForegroundColor Red
            Write-Host "   After installation, configure it with: rclone config" -ForegroundColor Yellow
            Write-Host "   Then set the remote name below (default: 'gdrive')" -ForegroundColor Yellow
            return $false
        }
        
        # Upload to Google Drive
        # Note: Replace 'gdrive' with your rclone remote name if different
        $remoteName = "gdrive"
        $remotePath = "$remoteName:/Service Board Imports"
        
        Write-Host "📤 Uploading $fileName to Google Drive..." -ForegroundColor Cyan
        $result = & rclone copy "$filePath" "$remotePath" --no-check-dest 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ Successfully uploaded $fileName" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ Failed to upload $fileName: $result" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "❌ Error uploading file: $_" -ForegroundColor Red
        return $false
    }
}

# Main monitoring loop
Write-Host "🚀 F2 Desktop Monitor started" -ForegroundColor Green
Write-Host "📁 Monitoring: $DESKTOP_PATH" -ForegroundColor Cyan
Write-Host "📂 Target folder: Service Board Imports (ID: $GOOGLE_DRIVE_FOLDER_ID)" -ForegroundColor Cyan
Write-Host "⏱️  Check interval: $CHECK_INTERVAL seconds" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Ctrl+C to stop monitoring" -ForegroundColor Yellow
Write-Host ""

$processedFiles = Get-ProcessedFiles
$lastCheck = @{}

while ($true) {
    try {
        # Get all Excel files on desktop
        $excelFiles = Get-ChildItem -Path $DESKTOP_PATH -Filter "*.xlsx" -ErrorAction SilentlyContinue
        
        foreach ($file in $excelFiles) {
            $fileName = $file.Name
            $filePath = $file.FullName
            
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
                Write-Host "⏳ File $fileName is still being written, waiting..." -ForegroundColor Yellow
                continue
            }
            
            # Upload file
            if (Upload-ToGoogleDrive -filePath $filePath -fileName $fileName) {
                # Mark as processed
                Add-ProcessedFile $fileName
                $processedFiles = Get-ProcessedFiles
                
                # Optionally move or delete the file from desktop
                # Uncomment one of these lines if you want to clean up:
                # Remove-Item $filePath -Force  # Delete from desktop
                # Move-Item $filePath "$DESKTOP_PATH\F2 Exports\$fileName" -Force  # Move to subfolder
            }
        }
    }
    catch {
        Write-Host "❌ Error in monitoring loop: $_" -ForegroundColor Red
    }
    
    # Wait before next check
    Start-Sleep -Seconds $CHECK_INTERVAL
}

