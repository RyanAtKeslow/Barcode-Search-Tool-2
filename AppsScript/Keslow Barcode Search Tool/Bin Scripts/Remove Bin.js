function removeBinAndUpdate() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
  
    var binColumn = 1; // Column A for bin numbers
    var lastRow = sheet.getLastRow();
  
    var selectedRange = sheet.getActiveRange();
    if (!selectedRange) {
      ui.alert("No cell selected. Please select a cell containing the bin number to remove.");
      return;
    }
  
    var inputBin = selectedRange.getValue().toString().trim();
  
    var binData = sheet.getRange(1, binColumn, lastRow, 1).getValues().flat(); // Read all bin numbers into an array

    var foundIndex = binData.indexOf(inputBin);
    if (foundIndex === -1) {
      ui.alert("Bin number not found.");
      return;
    }

    // Remove the bin by deleting the entire row
    sheet.deleteRow(foundIndex + 1); // +1 because sheet rows are 1-indexed

    // Remap the bin data after deletion
    binData = sheet.getRange(1, binColumn, lastRow - 1, 1).getValues().flat(); // Re-read all bin numbers into an array

    // Adjust subsequent bin numbers (-1) **WITHOUT CHANGING THE NEW BIN**
    var prefixMatch = inputBin.slice(0, 3); // Extract "4-E"
    var currentNumber = parseInt(inputBin.slice(3), 10);

    for (var i = foundIndex; i < binData.length; i++) {
      var match = binData[i].match(/^([0-9])-([A-Z])([0-9]{3})$/);
      if (match) {
        var prefix = match[1] + "-" + match[2];
        var number = parseInt(match[3], 10);
        if (prefix === prefixMatch && number > currentNumber && number.toString().charAt(0) === currentNumber.toString().charAt(0)) {
          number -= 1;
          binData[i] = prefix + number.toString().padStart(3, "0");
        }
      }
    }

    // Write updated bin numbers back to the sheet
    sheet.getRange(1, binColumn, binData.length, 1).setValues(binData.map(v => [v]));

    // Log the change
    var changeLogSheet = spreadsheet.getSheetByName("Change Log") || spreadsheet.insertSheet("Change Log");
    if (changeLogSheet.getLastRow() === 0) {
      changeLogSheet.appendRow(["Timestamp", "Change Description", "", "", "", "", "", "Status", "Col I", "Col J", "Col K", "Col L"]);
    }

    var timestamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd/yy HH:mm");
    var logEntryMessage = `Removed Bin: (${inputBin}). Subsequent bins adjusted accordingly.`;

    var logEntry = [timestamp, logEntryMessage, "", "", "", "", "", "", "", "", "", ""];
    changeLogSheet.appendRow(logEntry);

    var logRow = changeLogSheet.getLastRow();
    for (var col = 9; col <= 12; col++) { // Columns I to L
      var checkboxCell = changeLogSheet.getRange(logRow, col);
      var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      checkboxCell.setDataValidation(checkboxRule);
      checkboxCell.setValue(false);
    }

    ui.alert("Bin removed and changes logged successfully.");
  }