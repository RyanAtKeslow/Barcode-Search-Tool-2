/**
 * Adds a custom menu named "Custom Menu ESC" with items to run
 * findAvailableLens() and findAvailableCameras().
 * The menu appears each time the spreadsheet is opened or add-on installed.
 */
function onOpen(e) {
  createEscMenu_();
}

function onInstall(e) {
  createEscMenu_();
}

/**
 * Builds the custom menu and attaches it to the active spreadsheet UI.
 * Uses '_' suffix to avoid accidental exposure as a menu item.
 */
function createEscMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu ESC")
    .addItem("Find Available Lens", "findAvailableLens")
    .addItem("Find Available Cameras", "findAvailableCameras")
    .addToUi();
} 