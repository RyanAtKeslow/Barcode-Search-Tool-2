# Setting Up F2 Desktop Monitor to Run at Startup

## Quick Setup (Recommended)

1. **Run the setup script (NO ADMIN REQUIRED):**
   ```powershell
   cd "C:\Users\Ryan Griffith\Desktop\Barcode-Search-Tool-2"
   .\Setup-F2-Monitor-Startup.ps1
   ```
   **Note:** This creates a user-level task that runs for YOUR account. You do NOT need Administrator privileges!

2. **That's it!** The script will now run automatically every time you log in to Windows.

## How It Works

- The script runs as a **Windows Scheduled Task** that starts when you log in
- It runs **hidden in the background** - no PowerShell window will be visible
- It will continue running until you log out or shut down

## Manual Options

### Option 1: Startup Folder (Simple, but shows window)

1. Press `Win + R`, type `shell:startup`, press Enter
2. Create a shortcut to `F2-Desktop-Monitor.ps1`
3. Right-click the shortcut → Properties
4. Change "Target" to:
   ```
   powershell.exe -WindowStyle Minimized -ExecutionPolicy Bypass -File "C:\Users\Ryan Griffith\Desktop\Barcode-Search-Tool-2\F2-Desktop-Monitor.ps1"
   ```
5. Click OK

**Note:** This will show a minimized PowerShell window in the taskbar.

### Option 2: Scheduled Task (Hidden, like the setup script)

If you prefer to set it up manually:

1. Press `Win + R`, type `taskschd.msc`, press Enter
2. Click "Create Basic Task" in the right panel
3. Name: "F2 Desktop Monitor"
4. Trigger: "When I log on"
5. Action: "Start a program"
6. Program: `powershell.exe`
7. Arguments: `-WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Users\Ryan Griffith\Desktop\Barcode-Search-Tool-2\F2-Desktop-Monitor.ps1"`
8. Check "Open the Properties dialog for this task when I click Finish"
9. In Properties → General: Check "Run whether user is logged on or not" (optional)
10. Click OK

## Managing the Task

**Start the task manually:**
```powershell
Start-ScheduledTask -TaskName "F2 Desktop Monitor"
```

**Stop the task:**
```powershell
Stop-ScheduledTask -TaskName "F2 Desktop Monitor"
```

**Remove the startup task:**
```powershell
Unregister-ScheduledTask -TaskName "F2 Desktop Monitor" -Confirm:$false
```

**Check if it's running:**
```powershell
Get-ScheduledTask -TaskName "F2 Desktop Monitor" | Get-ScheduledTaskInfo
```

## Troubleshooting

**Task not starting?**
- Make sure PowerShell execution policy allows it: `Get-ExecutionPolicy` (should be RemoteSigned or Unrestricted)
- Check Task Scheduler for error messages
- Try running the script manually first to ensure it works

**Want to see the output?**
- The scheduled task runs hidden, but you can check the processed files log at: `%APPDATA%\F2DesktopMonitor\processed_files.txt`
- Or temporarily change `-WindowStyle Hidden` to `-WindowStyle Normal` to see the window

