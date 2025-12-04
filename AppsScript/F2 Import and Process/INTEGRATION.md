# F2 Import and Process - Integration Guide

## Menu Integration

The F2 Import menu needs to be added to the existing Custom Script Menu. 

### Option 1: Add to Existing Menu Script (Recommended)

In `Custom Script Menu (RTR Database).js`, add the F2 Import menu by calling `addF2ImportMenu(ui)`:

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
  
  // Add F2 Import menu
  addF2ImportMenu(ui);
}
```

### Option 2: Standalone Menu (If you prefer separate)

If you want the F2 Import menu to be completely separate, you can use the `onOpen()` function directly in `Custom Menu.js`, but you'll need to rename it to avoid conflicts, or use a different approach.

## Files Required

Make sure these files are in your Apps Script project:

1. `Process F2 Import.js` - Main processing script
2. `Custom Menu.js` - Menu function (call `addF2ImportMenu(ui)` from existing menu)

## Setup Steps

1. Copy `Process F2 Import.js` into your Apps Script project
2. Copy the `addF2ImportMenu()` function from `Custom Menu.js` into your existing menu script, OR
3. Add `addF2ImportMenu(ui);` call to your existing `onOpen()` function in `Custom Script Menu (RTR Database).js`

## Testing

After integration:
1. Open the spreadsheet
2. You should see a new "F2 Import" menu
3. Test with "View Processed Files" to verify it's working
4. Process a test Excel file from your Google Drive folder

