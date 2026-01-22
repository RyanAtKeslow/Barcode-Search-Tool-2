/**
 * Print Generic Bay - Document Generation Script
 * 
 * This script creates a formatted Google Doc with barcode bay data
 * for printing and physical record keeping.
 * 
 * Step-by-step process:
 * 1. Validates that the script is run from the "Barcode SEARCH" sheet
 * 2. Identifies the current user using email lookup with fallback logic
 * 3. Searches for username matches in row 2 of the sheet
 * 4. Handles multiple username matches with user selection dialog
 * 5. Prompts for username if not found or empty
 * 6. Extracts barcode bay name from the merged cell above username
 * 7. Determines column positions (barcode, item, bin) relative to username
 * 8. Collects all data from the three columns starting from row 4
 * 9. Filters out empty rows to create clean data set
 * 10. Creates timestamped Google Doc with formatted content
 * 11. Adds headers, bay name, and formatted table
 * 12. Saves document and moves to "Barcode Bay Printouts" folder
 * 13. Displays link to the created document
 * 
 * Document Structure:
 * - Header: Document creator and timestamp
 * - Subheader: Barcode bay name
 * - Table: Barcode | Item | Bin columns
 * - Filename: "MM/dd/yy HH:mm Username (BayName) Printout"
 * 
 * Features:
 * - User identification with multiple fallback methods
 * - Dynamic bay name detection
 * - Professional document formatting
 * - Automatic folder organization
 * - Print-ready output
 */
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
    Logger.log("🔄 Attempting to get user info from email...");
    const userInfo = fetchUserInfoFromEmail();
    username = userInfo.firstName;
    Logger.log(`✅ Got username from fetchUserInfoFromEmail: ${username}`);
  } catch (error) {
    Logger.log(`⚠️ fetchUserInfoFromEmail failed: ${error.toString()}`);
    try {
      Logger.log("🔄 Attempting fallback user identification...");
      const { nickname } = fetchUserEmailandNickname();
      username = nickname;
      Logger.log(`✅ Got username from fetchUserEmailandNickname: ${username}`);
    } catch (error) {
      Logger.log(`❌ Both user identification methods failed: ${error.toString()}`);
      SpreadsheetApp.getUi().alert("Could not identify user. Please ensure you are logged in with your company email.");
      return;
    }
  }
  
  // Find all username matches in row 2
  // Use a reasonable maximum column limit to prevent accessing cells beyond actual data
  const maxReasonableColumn = 100; // Limit search to first 100 columns (A through CV)
  const lastColumn = Math.min(activeSheet.getLastColumn(), maxReasonableColumn);
  Logger.log(`🔍 Searching row 2 from column 1 to ${lastColumn} for username: ${username}`);
  
  var row2Range = activeSheet.getRange(2, 1, 1, lastColumn);
  var row2Values = row2Range.getValues()[0];
  var usernameMatches = [];
  
  for (let j = 0; j < row2Values.length; j++) {
    if (row2Values[j] && row2Values[j].toString().toLowerCase().includes(username.toLowerCase())) {
      const cellA1 = activeSheet.getRange(2, j + 1).getA1Notation();
      Logger.log(`✅ Found username match at ${cellA1}: ${row2Values[j]}`);
      usernameMatches.push({
        cellA1: cellA1,
        value: row2Values[j]
      });
    }
  }
  
  Logger.log(`📊 Found ${usernameMatches.length} username match(es)`);
  
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
  
  // Validate cell reference before using it
  Logger.log(`🔍 Processing selected cell: ${selectedCellA1}`);
  
  // Check if cell reference is reasonable (not beyond column 100)
  const cellMatch = selectedCellA1.match(/^([A-Z]+)(\d+)$/);
  if (cellMatch) {
    const columnLetters = cellMatch[1];
    // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
    let columnNum = 0;
    for (let i = 0; i < columnLetters.length; i++) {
      columnNum = columnNum * 26 + (columnLetters.charCodeAt(i) - 64);
    }
    if (columnNum > 100) {
      Logger.log(`❌ Invalid cell reference ${selectedCellA1} - column ${columnNum} is beyond reasonable range`);
      SpreadsheetApp.getUi().alert(`Error: Invalid cell reference ${selectedCellA1}. Please contact support.`);
      return;
    }
  }
  
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var userName = usernameCell.getValue();
  // Handle different value types (string, number, Date, etc.)
  if (userName !== null && userName !== undefined) {
    userName = userName.toString().trim();
  } else {
    userName = "";
  }
  Logger.log(`📝 Username from cell ${selectedCellA1}: "${userName}"`);

  // Dynamically get the barcode bay name from the merged cell above the username cell
  var userRow = usernameCell.getRow();
  var userColumn = usernameCell.getColumn();
  var barcodeBayName = activeSheet.getRange(userRow - 1, userColumn).getValue().toString().trim(); 

  // Function to capitalize words
  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function(char) { return char.toUpperCase(); });
  }

  // Ensure username is properly formatted if it exists
  if (userName !== "") {
    const capitalizedName = capitalizeWords(userName);
    // Only update the cell if the capitalization changed (to avoid unnecessary writes)
    if (capitalizedName !== userName) {
      Logger.log(`📝 Capitalizing username: "${userName}" -> "${capitalizedName}"`);
      usernameCell.setValue(capitalizedName);
      userName = capitalizedName;
    }
  }

  // Prompt user if no name is entered
  if (userName === "") {
    Logger.log(`⚠️ Username cell ${selectedCellA1} is empty, prompting user to enter name`);
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