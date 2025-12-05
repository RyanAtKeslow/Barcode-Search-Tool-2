/**
 * NOTE: onOpen() has been moved to Custom Script Menu (RTR Database).js to consolidate all menus
 * and avoid conflicts. This function is commented out but kept for reference.
 * 
 * The ESC menu is now created in the main Custom Script Menu (RTR Database).js file.
 */
/*
function onOpen(e) {
  createEscMenu_();
}

function onInstall(e) {
  createEscMenu_();
}

function createEscMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu ESC")
    .addItem("Find Available Lens", "findAvailableLens")
    .addItem("Find Available Cameras", "findAvailableCameras")
    .addToUi();
}
*/ 