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
    .addToUi();
  
  // Create Prep Bay menu
  ui.createMenu('Prep Bay')
    .addItem('Test Prep Bay Refresh', 'testPrepBayRefresh')
    .addSeparator()
    .addItem('Clear All Prep Bays', 'clearAllPrepBays')
    .addToUi();
  
  // Create Custom Menu ESC (commented out)
  // ui.createMenu('Custom Menu ESC')
  //   .addItem('Find Available Lens', 'findAvailableLens')
  //   .addItem('Find Available Cameras', 'findAvailableCameras')
  //   .addToUi();
} 