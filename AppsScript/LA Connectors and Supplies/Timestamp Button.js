/**
* Sets the current timestamp in a selected cell in Column L.
* It then automatically extracts the user's name from their email
* and places it in the adjacent cell in Column M.
* Displays an error if the selection is outside of Column L.
*/
function setNow() {
  const TARGET_COLUMN = 12; // Column L

  const range = SpreadsheetApp.getActiveRange();

  if (range.getColumn() !== TARGET_COLUMN || range.getLastColumn() !== TARGET_COLUMN) {
    SpreadsheetApp.getUi().alert('Action Canceled: Please select one or more cells in Column L only.');
    return;
  }

  range.setValue(new Date());

  const currentUserEmail = Session.getActiveUser().getEmail();
  if (currentUserEmail) {
    const namePart = currentUserEmail.split('@')[0];
    const signature = namePart.charAt(0).toUpperCase() + namePart.slice(1);
    range.offset(0, 1).setValue(signature);
  }
}
