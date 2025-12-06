# F2 Desktop Monitor - Setup Guide

This script automatically monitors your Desktop for F2 Service Board export files and uploads them to your Google Drive "Service Board Imports" folder.

## Quick Start - Choose Your Method

### Option 1: Simple Method (Recommended) - Using Google Drive for Desktop

**Easiest setup - no API credentials needed!**

1. **Install Google Drive for Desktop**
   - Download: https://www.google.com/drive/download/
   - Install and sign in with your Google account

2. **Set up sync folder**
   - In Google Drive for Desktop, add a folder to sync
   - Point it to sync your "Service Board Imports" folder
   - Note the local path (usually `G:\My Drive\Service Board Imports`)

3. **Update the simple script**
   - Open `F2 Desktop Monitor (Simple).ps1`
   - Update `$SYNC_FOLDER_PATH` to your local Google Drive folder path

4. **Run the script**
   ```powershell
   .\F2 Desktop Monitor (Simple).ps1
   ```

That's it! The script moves files from Desktop to the sync folder, and Google Drive for Desktop automatically uploads them.

### Option 2: Using rclone (More Control)

1. **Install rclone**
   ```powershell
   # Using winget (Windows 10/11)
   winget install rclone.rclone
   
   # Or download from https://rclone.org/downloads/
   ```

2. **Configure rclone for Google Drive**
   ```powershell
   rclone config
   ```
   
   Follow the prompts:
   - Choose "n" for new remote
   - Name it `gdrive` (or whatever you prefer)
   - Choose "Google Drive" (option 15)
   - Follow the OAuth setup (it will open a browser)
   - Grant permissions to access Google Drive
   - Leave other options as default

3. **Test the connection**
   ```powershell
   rclone lsd gdrive:
   ```
   This should list your Google Drive folders.

4. **Update the script** (if you used a different remote name)
   - Open `F2 Desktop Monitor.ps1`
   - Change `$remoteName = "gdrive"` to your remote name

5. **Run the script**
   ```powershell
   .\F2 Desktop Monitor.ps1
   ```

### Option 2: Using Google Drive for Desktop (Alternative)

If you prefer not to use rclone, you can use Google Drive for Desktop:

1. **Install Google Drive for Desktop**
   - Download: https://www.google.com/drive/download/
   - Install and sign in

2. **Set up a sync folder**
   - Create a folder on your Desktop called "F2 Exports"
   - In Google Drive for Desktop settings, add this folder to sync
   - Configure it to sync to your "Service Board Imports" folder

3. **Use a simpler PowerShell script** (just moves files to the sync folder)
   - This would be a simpler script that just moves files from Desktop to the sync folder
   - Google Drive for Desktop handles the upload automatically

## Running the Script

### Manual Run
```powershell
.\F2 Desktop Monitor.ps1
```

### Run at Startup (Windows)

1. **Create a scheduled task:**
   - Press `Win + R`, type `taskschd.msc`
   - Create Basic Task
   - Name: "F2 Desktop Monitor"
   - Trigger: "When I log on"
   - Action: "Start a program"
   - Program: `powershell.exe`
   - Arguments: `-WindowStyle Hidden -File "C:\path\to\F2 Desktop Monitor.ps1"`

2. **Or add to Startup folder:**
   - Press `Win + R`, type `shell:startup`
   - Create a shortcut to the PowerShell script

### Run as Background Service

You can also run it as a Windows service using NSSM (Non-Sucking Service Manager):
```powershell
# Download NSSM from https://nssm.cc/download
# Install the service
nssm install F2DesktopMonitor "powershell.exe" "-File C:\path\to\F2 Desktop Monitor.ps1"
nssm start F2DesktopMonitor
```

## Configuration

Edit `F2 Desktop Monitor.ps1` to customize:

- `$GOOGLE_DRIVE_FOLDER_ID`: Your Google Drive folder ID (already set)
- `$DESKTOP_PATH`: Path to monitor (default: your Desktop)
- `$CHECK_INTERVAL`: How often to check for new files (default: 30 seconds)
- `$PROCESSED_FILES_LOG`: Where to store the log of processed files

## Troubleshooting

### rclone not found
- Make sure rclone is in your PATH
- Or use full path: `C:\path\to\rclone.exe`

### Authentication issues
- Re-run `rclone config` and re-authenticate
- Check that you granted the correct permissions

### Files not uploading
- Check rclone connection: `rclone lsd gdrive:`
- Verify folder ID is correct
- Check the processed files log to see if files are being detected

### Script stops running
- Check PowerShell execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
- Run as administrator if needed

## Notes

- The script keeps a log of processed files to avoid re-uploading
- Files are checked every 30 seconds (configurable)
- The script waits 5 seconds after file modification to ensure it's fully written
- By default, files remain on your Desktop after upload (you can modify the script to delete or move them)

