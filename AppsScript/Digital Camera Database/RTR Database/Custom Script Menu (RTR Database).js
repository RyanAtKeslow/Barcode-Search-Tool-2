/**
 * Creates all custom menus when the spreadsheet is opened
 * Combines menus from multiple scripts to avoid conflicts
 */
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