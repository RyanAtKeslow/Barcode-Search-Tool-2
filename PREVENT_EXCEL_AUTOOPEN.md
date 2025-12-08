# How to Prevent Excel from Auto-Opening Files

When F2 exports a file, Windows may automatically open it in Excel. This can cause the F2 Desktop Monitor to fail when trying to move the file because Excel has it locked.

## Solution 1: Change Windows File Association (Recommended)

This prevents Excel from auto-opening `.xlsx` files when they're created:

1. **Open File Explorer** and navigate to any folder
2. **Right-click** on any `.xlsx` file
3. Select **"Open with"** → **"Choose another app"**
4. **Uncheck** "Always use this app to open .xlsx files" (if checked)
5. Select **"Notepad"** or any non-Excel app (just for this step)
6. Click **OK**
7. **Right-click** the file again → **"Open with"** → **"Choose another app"**
8. **Check** "Always use this app to open .xlsx files"
9. Select **"Excel"** (or your preferred Excel app)
10. Click **OK**

**Note:** This changes the default app, but F2 may still trigger Excel to open. See Solution 2.

## Solution 2: Disable Auto-Open in F2 (If Available)

Check F2's export settings:
- Look for options like "Open file after export" or "Auto-open exported files"
- Disable this setting if available
- The exact location depends on your F2 version

## Solution 3: Use the Improved Monitor Script

The updated `F2-Desktop-Monitor.ps1` script now:
- ✅ Waits longer (15 seconds) after file creation before attempting to move
- ✅ Detects when files are locked by Excel
- ✅ Automatically retries up to 6 times with increasing wait times
- ✅ Logs helpful messages when files are locked

The script will automatically handle files that Excel has open and will move them once Excel releases the lock (when you close Excel or the file).

## Solution 4: Close Excel Prompt Automatically (Advanced)

If you want to automatically close Excel when it opens (since you mentioned it's not activated), you could:

1. Create a small script that monitors for Excel windows and closes them
2. Or use a tool like AutoHotkey to automatically close Excel activation prompts

However, the improved monitor script should handle this automatically now, so you may not need this.

## Testing

After making changes:
1. Export a file from F2
2. Watch the monitor console/log - it should wait and retry if Excel has the file open
3. Close Excel (or the file in Excel)
4. The monitor should successfully move the file on the next retry

## Monitor Log Location

Check the monitor log for details:
```
%APPDATA%\F2DesktopMonitor\monitor.log
```

The log will show retry attempts and when files are successfully moved.

