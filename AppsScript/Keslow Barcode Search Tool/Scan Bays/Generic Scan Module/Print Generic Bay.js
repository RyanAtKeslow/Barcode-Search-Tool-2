function printGenericBay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Barcode SEARCH") {
    SpreadsheetApp.getUi().alert("Please run this script from the 'Barcode SEARCH' sheet");
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
    continuePrintGenericBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    continuePrintGenericBay(usernameMatches[0].cellA1);
  } else {
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continuePrintGenericBay(selectedCellA1) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var userName = usernameCell.getValue().toString().trim();

  // Dynamically get the barcode bay name from the merged cell above the username cell
  var userRow = usernameCell.getRow();
  var userColumn = usernameCell.getColumn();
  var barcodeBayName = activeSheet.getRange(userRow - 1, userColumn).getValue().toString().trim(); 

  // Function to capitalize words
  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function(char) { return char.toUpperCase(); });
  }

  // Ensure username is properly formatted
  if (userName !== "") {
    userName = capitalizeWords(userName);
    usernameCell.setValue(userName);
  }

  // Prompt user if no name is entered
  if (userName === "") {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Please Enter Your Name", 
      "(Please sign your work by putting your name on the Barcode Bay)", 
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      ui.alert("Please sign your work to continue");
      return;
    }

    if (response.getResponseText().trim() !== "") {
      userName = capitalizeWords(response.getResponseText().trim());
      usernameCell.setValue(userName);
    } else {
      ui.alert("You must sign your work to continue");
      return;
    }
  }

  // Dynamically determine column positions
  var barcodeColumn = userColumn - 1;  // Left of username
  var itemColumn = userColumn;         // Same as username
  var binColumn = userColumn + 1;      // Right of username

  // Get the values from respective columns, starting from row 4
  var numRows = activeSheet.getLastRow() - 3; // Since we start from row 4
  var barcodeRange = activeSheet.getRange(4, barcodeColumn, numRows, 1).getValues();
  var itemRange = activeSheet.getRange(4, itemColumn, numRows, 1).getValues();
  var binRange = activeSheet.getRange(4, binColumn, numRows, 1).getValues();

  // Store formatted values
  var formattedValues = [];
  for (var i = 0; i < numRows; i++) {
    var barcode = barcodeRange[i][0];
    var itemName = itemRange[i][0];
    var bin = binRange[i][0];

    if (barcode !== "" || itemName !== "" || bin !== "") {
      formattedValues.push([String(barcode), itemName, bin]);
    }
  }

  // Get current timestamp
  var timeZone = "America/Los_Angeles";
  var date = Utilities.formatDate(new Date(), timeZone, "MM/dd/yy HH:mm");

  // Document naming convention with dynamic barcode bay
  var docName = date + " " + userName + " (" + barcodeBayName + ") Printout";

  // Create a new Google Doc
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // Add headers
  body.appendParagraph("Document Created by: " + userName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Timestamp: " + date).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  // Insert barcode bay name as a subheading before the table
  body.appendParagraph("Barcode Bay: " + barcodeBayName).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Create a table
  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell('Barcode');
  headerRow.appendTableCell('Item');
  headerRow.appendTableCell('Bin');

  // Add table rows
  formattedValues.forEach(function(row) {
    var tableRow = table.appendTableRow();
    row.forEach(function(cell) {
      tableRow.appendTableCell(cell);
    });
  });

  // Save and close document
  doc.saveAndClose();
  var docUrl = doc.getUrl();

  // Display a link to the document
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align: center;">' +
    'Your document is ready for printing.<br><br>' +
    '<a href="' + docUrl + '" target="_blank" style="font-size: 16px; color: blue; text-decoration: none;">Click here to open and print your document</a>' +
    '</div>'
  )
  .setWidth(400)
  .setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Document Created');

  // Move document to "Barcode Bay Printouts" folder
  var rootFolder = DriveApp.getRootFolder();
  var folderIterator = rootFolder.getFoldersByName("Barcode Bay Printouts");
  var printoutsFolder = folderIterator.hasNext() ? folderIterator.next() : rootFolder.createFolder("Barcode Bay Printouts");

  var docFile = DriveApp.getFileById(doc.getId());
  printoutsFolder.addFile(docFile);
  rootFolder.removeFile(docFile);
} 