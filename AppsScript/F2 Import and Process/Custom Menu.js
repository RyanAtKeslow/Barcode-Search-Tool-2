/**
 * Custom Menu for F2 Import and Process
 * 
 * This function should be called from the existing Custom Script Menu (RTR Database).js
 * Add this line to the onOpen() function in that file:
 *   addF2ImportMenu(ui);
 */
function addF2ImportMenu(ui) {
  // Create "F2 Import" menu
  ui.createMenu('F2 Import')
    .addItem('Process Imports', 'processF2Imports')
    .addSeparator()
    .addItem('View Processed Files', 'getProcessedFilesSummary')
    .addItem('Reset Processed Files List', 'resetProcessedFilesList')
    .addToUi();
}

