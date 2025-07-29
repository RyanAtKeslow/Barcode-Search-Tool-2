function addBin() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var binColumn = 1; // Column A (Bin Numbers)

  // Step 1: Ask the user for the new bin number
  var userInput = Browser.inputBox(
    "Enter new bin number to add",
    "Please enter the new bin number (e.g., S-A100 or 3-A402):",
    Browser.Buttons.OK_CANCEL
  );

  if (userInput === "cancel") {
    ui.alert("Input canceled.");
    return;
  }

  var capitalizedInput = userInput.toUpperCase();

  // Step 2: Validate the bin number format
  if (!/^[0-9]-[A-Z][0-9]{3}$/i.test(capitalizedInput)) {
    ui.alert("Invalid input. Please provide a valid bin number (e.g., S-A100 or 3-A402).");
    return;
  }

  // Step 3: Ask for the bin's name
  var inputBValue = Browser.inputBox(
    "Enter the bin's name",
    "Copy and Paste from Flawless",
    Browser.Buttons.OK_CANCEL
  );

  if (inputBValue === "cancel" || !inputBValue.trim()) {
    ui.alert("Invalid or canceled input. Please provide a valid bin name.");
    return;
  }

  // Step 4: Read all bin numbers into an array
  var binData = sheet.getRange(2, binColumn, lastRow - 1, 1).getValues().flat(); // Read column A (excluding header)

  // Step 5: Find the first bin that matches the input's prefix (e.g., 4-E)
  var prefixMatch = capitalizedInput.slice(0, 3); // Extract "4-E"
  var insertIndex = binData.findIndex(bin => bin.startsWith(prefixMatch) && bin === capitalizedInput);

  if (insertIndex === -1) {
    // If no exact match is found, find the bin with a number one less than the input
    var numberMatch = parseInt(capitalizedInput.slice(3), 10) - 1;
    insertIndex = binData.findIndex(bin => bin.startsWith(prefixMatch) && parseInt(bin.slice(3), 10) === numberMatch);
    if (insertIndex !== -1) {
      insertIndex += 1; // Insert below the found bin
    } else {
      insertIndex = binData.length; // Append at the end if no match is found
    }
  }

  var actualInsertRow = insertIndex + 2; // Adjust for header row

  // Step 6: Insert a new row above the found row
  sheet.insertRowBefore(actualInsertRow);

  // Step 7: Set bin number and name in the new row
  sheet.getRange(actualInsertRow, binColumn).setValue(capitalizedInput);
  sheet.getRange(actualInsertRow, 2).setValue(inputBValue);

  // Step 8: Adjust subsequent bin numbers (+1) **WITHOUT CHANGING THE NEW BIN**
  for (var i = insertIndex; i < binData.length; i++) {
    var match = binData[i].match(/^([0-9])-([A-Z])([0-9]{3})$/);
    if (match) {
      var prefix = match[1] + "-" + match[2];
      var currentNumber = parseInt(match[3], 10);
      var newNumber = (currentNumber + 1).toString().padStart(3, "0");
      if (newNumber.charAt(0) !== match[3].charAt(0) || match[3].charAt(0) !== capitalizedInput.charAt(3)) {
        break; // Stop if the hundredth place changes or if it's a different series
      }
      binData[i] = prefix + newNumber;
    }
  }

  // Step 9: Write updated bin numbers back to the sheet (excluding the newly added bin)
  sheet.getRange(actualInsertRow + 1, binColumn, binData.length - insertIndex, 1).setValues(binData.slice(insertIndex).map(v => [v]));

  // Step 11: Insert an empty checkbox in column E
  sheet.getRange(actualInsertRow, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sheet.getRange(actualInsertRow, 5).setValue(false);

  // Step 12: Log the change in "Change Log" sheet
  var changeLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Change Log") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Change Log");
  var timestamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd/yy HH:mm");
  var changeMessage = `New Bin: (${capitalizedInput}) added at row ${actualInsertRow}. Subsequent bins adjusted accordingly.`;

  changeLogSheet.appendRow([timestamp, changeMessage]);

  // Step 13: Add checkboxes in columns I–L in "Change Log"
  var logRow = changeLogSheet.getLastRow();
  var checkboxRange = changeLogSheet.getRange(logRow, 9, 1, 4);
  var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange.setDataValidation(checkboxRule);
  checkboxRange.setValue(false);

  ui.alert("Bin added successfully!");
}