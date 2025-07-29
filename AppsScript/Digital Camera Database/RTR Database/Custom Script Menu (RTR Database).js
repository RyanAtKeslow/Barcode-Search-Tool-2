function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create "Camera Forecast" menu
  ui.createMenu('Camera Forecast')
    .addItem('Generate Forecast', 'getCameraForecast')
    .addToUi();
} 