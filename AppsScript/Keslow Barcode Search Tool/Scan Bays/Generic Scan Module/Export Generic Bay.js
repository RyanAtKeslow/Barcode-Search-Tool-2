/**
 * Export Generic Bay - CSV Export Script
 * 
 * This script exports barcode data from a selected bay to a CSV file
 * for external use or backup purposes.
 * 
 * Step-by-step process:
 * 1. Validates that the script is run from the "Barcode SEARCH" sheet
 * 2. Identifies the current user using email lookup with fallback logic
 * 3. Searches for username matches in row 2 of the sheet
 * 4. Handles multiple username matches with user selection dialog
 * 5. Prompts for username if not found or empty
 * 6. Capitalizes and formats the username properly
 * 7. Determines the barcode column (left of username column)
 * 8. Extracts all barcode values from the bay starting from row 4
 * 9. Validates that barcodes exist before proceeding
 * 10. Creates CSV content with one barcode per line
 * 11. Generates timestamped filename with username
 * 12. Creates or finds "CSV Exports" folder in Google Drive
 * 13. Saves CSV file and displays download link
 * 
 * Features:
 * - User identification with multiple fallback methods
 * - Multiple username match handling
 * - Automatic file naming with timestamps
 * - Google Drive integration
 * - User-friendly download interface
 * 
 * Output: CSV file with format: "MM-dd-yy_HH-mm_Username_Barcodes.csv"
 */
function exportGenericBay() {
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
    continueExportGenericBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    continueExportGenericBay(usernameMatches[0].cellA1);
  } else {
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueExportGenericBay(selectedCellA1) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var userName = usernameCell.getValue().toString().trim();
  
  // Function to capitalize each word
  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function(char) { return char.toUpperCase(); });
  }

  // Ensure username is properly formatted
  if (userName !== "") {
    userName = capitalizeWords(userName);
    usernameCell.setValue(userName); // Ensure capitalization in the sheet
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

  // Dynamically determine the barcode column (left of username cell)
  var userColumn = usernameCell.getColumn(); 
  var barcodeColumn = userColumn - 1; 
  var barcodeRange = activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1); // Start from row 4

  // Get the barcode values
  var barcodes = barcodeRange.getValues().flat().filter(String);

  // Ensure there's data to export
  if (barcodes.length === 0) {
    SpreadsheetApp.getUi().alert("No barcodes found in column " + String.fromCharCode(64 + barcodeColumn) + ".");
    return;
  }

  // Convert the barcodes array to a CSV string (one per line)
  var csvContent = barcodes.join("\n");

  // Get current timestamp
  var date = getConsistentTimestamp();

  // Define file name
  var fileName = date + "_" + userName + "_Barcodes.csv";

  // Save to Shared Drive
  Logger.log("🔄 Starting Shared Drive save...");
  saveToSharedDrive('CSV Exports', fileName, csvContent);
  Logger.log("✅ Shared Drive save completed");
  
  // Also save to user's personal Drive as backup
  Logger.log("🔄 Starting personal Drive save...");
  try {
    var rootFolder = DriveApp.getRootFolder();
    Logger.log("✅ Got root folder");
    
    var folderIterator = rootFolder.getFoldersByName("CSV Exports");
    Logger.log("✅ Got folder iterator");
    
    var csvFolder = folderIterator.hasNext() ? folderIterator.next() : rootFolder.createFolder("CSV Exports");
    Logger.log(`✅ Got/created CSV folder: ${csvFolder.getName()}`);

    // Create the CSV file in Drive
    Logger.log("🔄 Creating CSV file...");
    var csvFile = csvFolder.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
    Logger.log(`✅ CSV file created: ${csvFile.getName()} (ID: ${csvFile.getId()})`);
  } catch (error) {
    Logger.log(`❌ Error in personal Drive save: ${error.toString()}`);
    Logger.log(`Error details: ${JSON.stringify(error)}`);
    throw error; // Re-throw to see the full error
  }

  // Generate the file link (standard Google Drive link)
  var fileUrl = csvFile.getUrl();

  // Display the file link in a dialog
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align: center;">' +
    '<p>Your CSV file has been created in Google Drive.</p>' +
    '<a href="' + fileUrl + '" target="_blank" style="font-size: 16px; color: blue; text-decoration: none;">Click here to open the file</a>' +
    '<p style="font-size: 12px; color: #666; margin-top: 10px;">Once opened, you can download it using the download button in Google Drive.</p>' +
    '</div>'
  )
  .setWidth(450)
  .setHeight(140);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'CSV File Created');
}

// saveToSharedDrive function is now centralized in Data Model.js

function setSelectedUsernameCell(cellA1) {
  PropertiesService.getScriptProperties().setProperty('selectedUsernameCell', cellA1);
} 