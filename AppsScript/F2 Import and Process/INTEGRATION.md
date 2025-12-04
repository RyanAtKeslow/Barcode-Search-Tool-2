# F2 Import and Process - Integration Guide

## Menu Integration

The F2 Import menu has been integrated directly into `Custom Script Menu (RTR Database).js`. 

The menu is now part of the main menu file:

```javascript
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  // Create Serial Database menu
  ui.createMenu('Serial Database')
    .addItem('Refresh Serial Database', 'copySerialDatabase')
    .addToUi();
  
  // Create Camera Forecast menu
  ui.createMenu('Camera Forecast')
    .addItem('Generate Forecast', 'getCameraForecast')
    .addToUi();
  
  // Create F2 Import menu
  ui.createMenu('F2 Import')
    .addItem('Process Imports', 'processF2Imports')
    .addSeparator()
    .addItem('View Processed Files', 'getProcessedFilesSummary')
    .addItem('Reset Processed Files List', 'resetProcessedFilesList')
    .addToUi();
}
```

## Files Required

Make sure these files are in your Apps Script project:

1. `Process F2 Import.js` - Main processing script
2. `Custom Script Menu (RTR Database).js` - Updated with F2 Import menu (already integrated)

**Note:** `Custom Menu.js` is no longer needed - the menu has been integrated directly into the main menu file.

## Testing

After integration:
1. Open the spreadsheet
2. You should see a new "F2 Import" menu
3. Test with "View Processed Files" to verify it's working
4. Process a test Excel file from your Google Drive folder

