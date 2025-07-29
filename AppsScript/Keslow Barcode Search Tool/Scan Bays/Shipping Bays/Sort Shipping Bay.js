function sortShippingBay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Receiving Bays") {
    SpreadsheetApp.getUi().alert("Please run this script from the 'Receiving Bays' sheet");
    return;
  }

  // Try to get user info with fallback logic
  let username;
  try {
    const userInfo = fetchUserInfoFromEmail();
    username = userInfo.firstName;
  } catch (error) {
    try {
      const { nickname } = fetchUserEmailandNickname();
      username = nickname;
    } catch (error) {
      SpreadsheetApp.getUi().alert("Could not identify user. Please ensure you are logged in with your company email.");
      return;
    }
  }
  
  // Find all username matches in row 2
  var row2Range = activeSheet.getRange(2, 1, 1, activeSheet.getLastColumn());
  var row2Values = row2Range.getValues()[0];
  var usernameMatches = [];
  
  for (let j = 0; j < row2Values.length; j++) {
    if (row2Values[j] && row2Values[j].toString().toLowerCase().includes(username.toLowerCase())) {
      usernameMatches.push({
        cellA1: activeSheet.getRange(2, j + 1).getA1Notation(),
        value: row2Values[j]
      });
    }
  }
  
  // Handle multiple matches
  if (usernameMatches.length > 1) {
    const selectedMatch = setSelectedMatch(usernameMatches);
    continueSortShippingBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    continueSortShippingBay(usernameMatches[0].cellA1);
  } else {
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueSortShippingBay(selectedCellA1) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var userName = usernameCell.getValue().toString().trim();
  Logger.log("User name: " + userName); // Log the user name

  // Function to capitalize each word in the user's name
  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function(char) {
      return char.toUpperCase();
    });
  }

  // If B2 is empty, prompt the user for their name
  if (userName === "") {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Please Enter Your Name", 
      "(Please sign the scanner's name before sorting the bay)", 
      ui.ButtonSet.OK_CANCEL
    );

    // If the user clicks cancel, exit the script
    if (response.getSelectedButton() == ui.Button.CANCEL) {
      ui.alert("Please sign your work to continue");
      return; // Exit the script
    }

    // If the user provides a name, insert it into B2
    if (response.getResponseText().trim() !== "") {
      userName = capitalizeWords(response.getResponseText().trim());
      usernameCell.setValue(userName); // Store the formatted username in B2
    } else {
      ui.alert("Please sign your work to continue");
      return; // Exit the script if no name is entered
    }
  } else {
    // If B2 is not empty, ensure each word is capitalized
    userName = capitalizeWords(userName);
    usernameCell.setValue(userName); // Update B2 with the formatted name
  }

  // Dynamically determine column positions
  var userRow = usernameCell.getRow();
  var userColumn = usernameCell.getColumn();
  var barcodeColumn = userColumn - 1;  // Barcode is to the left of the username column
  var itemColumn = userColumn;         // Item is in the same column as the username
  var binColumn = userColumn + 1;      // Bin is to the right of the username column

  // Get the barcode, item, and bin number data starting from row 4
  var range = activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 3); // Starting from row 4
  var values = range.getValues();
  Logger.log("Original data: " + JSON.stringify(values)); // Log the original values

  // Filter out rows where barcode (column A) or bin number (column C) is empty
  values = values.filter(function(row) {
    return row[0] !== "" && row[2] !== ""; // Only keep rows with both barcode and bin number
  });
  Logger.log("Filtered data: " + JSON.stringify(values)); // Log the filtered values

  // Pair barcodes (column A) with bin numbers (column C)
  var barcodePairs = values.map(row => [row[0], row[2]]); // [Barcode, Bin Number]

  // Separate items with "No Bin" from others
  var noBinItems = barcodePairs.filter(row => row[1].toString().trim() === "No Bin");
  var sortedItems = barcodePairs.filter(row => row[1].toString().trim() !== "No Bin");

  // Sort valid bin entries alphanumerically
  sortedItems.sort((a, b) => a[1].localeCompare(b[1], undefined, { numeric: true }));

  // Merge sorted items first, followed by "No Bin" items at the bottom
  var finalOrder = sortedItems.concat(noBinItems).map(row => [row[0]]); // Extract only barcodes

  // Clear all barcode cells in the target column and rewrite the barcodes without gaps
  var totalRows = finalOrder.length;
  var startRow = 4;  // Start from row 4
  var rangeToClear = activeSheet.getRange(startRow, barcodeColumn, activeSheet.getLastRow() - 3, 1);
  rangeToClear.clearContent();

  // Write the sorted barcodes to the sheet without gaps
  activeSheet.getRange(startRow, barcodeColumn, totalRows, 1).setValues(finalOrder);

  Logger.log("Sorted barcodes written back to column A");

  // Ensure all changes are applied immediately
  SpreadsheetApp.flush();
  Logger.log("Flush complete"); // Log after flush
}