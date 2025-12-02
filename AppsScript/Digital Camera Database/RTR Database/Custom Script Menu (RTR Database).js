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
} 