/**
 * Reset Generic Bay - Bay Reset and Data Processing Script
 * 
 * This script resets a barcode bay by processing items based on their bin status,
 * organizing them into appropriate tracking sheets, and updating analytics.
 * 
 * Step-by-step process:
 * 1. Validates that the script is run from the "Barcode SEARCH" sheet
 * 2. Identifies the current user using email lookup with fallback logic
 * 3. Searches for username matches in row 2 of the sheet
 * 4. Handles multiple username matches with user selection dialog
 * 5. Prompts for username if not found or empty
 * 6. Determines column positions (barcode, item, bin) relative to username
 * 7. Processes items based on bin status:
 *    - "No Bin" items → HOMELESS GEAR sheet (with quantity tracking)
 *    - "LOST/DISPOSED/INACTIVE" items → Lost & Found sheet
 * 8. Cleans item names by removing status keywords
 * 9. Extracts consigner information from item names
 * 10. Updates analytics counters for barcodes and Lost & Found items
 * 11. Saves barcode data to CSV archive
 * 12. Clears the bay data and username cell
 * 
 * Data Processing:
 * - Homeless gear: Items without assigned bins, tracked by quantity
 * - Lost & Found: Items with status keywords, includes consigner info
 * - Analytics: Updates total barcode count and Lost & Found count
 * - Archive: Saves all processed barcodes to timestamped CSV files
 * 
 * Features:
 * - Duplicate prevention in Lost & Found
 * - Quantity tracking for homeless gear
 * - Consigner extraction and default assignment
 * - Comprehensive analytics updates
 * - Data archiving and cleanup
 */

function resetGenericBay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Barcode SEARCH") {
    SpreadsheetApp.getUi().alert("Please run this script from the 'Barcode SEARCH' sheet");
    return;
  }
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");

  // Try to get user info with fallback logic
  let userFirstName;
  try {
    const userInfo = fetchUserInfoFromEmail();
    userFirstName = userInfo.firstName;
  } catch (error) {
    try {
      const { nickname } = fetchUserEmailandNickname();
      userFirstName = nickname;
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
    if (row2Values[j] && row2Values[j].toString().toLowerCase().includes(userFirstName.toLowerCase())) {
      usernameMatches.push({
        cellA1: activeSheet.getRange(2, j + 1).getA1Notation(),
        value: row2Values[j]
      });
    }
  }
  
  // Handle multiple matches
  if (usernameMatches.length > 1) {
    const selectedMatch = setSelectedMatch(usernameMatches);
    continueResetGenericBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    continueResetGenericBay(usernameMatches[0].cellA1);
  } else {
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueResetGenericBay(selectedCellA1) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");
  
  // Get existing barcodes in Lost & Found to avoid duplicates
  var existingLFBarcodes = new Set(
    lostAndFoundSheet.getRange("A:A")
      .getValues()
      .flat()
      .filter(String)
  );
  
  var jobInfo = usernameCell.getValue().toString().trim();

  var userColumn = usernameCell.getColumn();
  var barcodeColumn = userColumn - 1;
  var binsColumn = userColumn + 1;

  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function (char) {
      return char.toUpperCase();
    });
  }

  if (jobInfo === "") {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Please Enter Your Name",
      "(Please sign the scanner's name before resetting the bay)",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      ui.alert("Please sign your work to continue");
      return;
    }

    if (response.getResponseText().trim() !== "") {
      jobInfo = capitalizeWords(response.getResponseText().trim());
      usernameCell.setValue(jobInfo);
    } else {
      ui.alert("You must sign your work to continue");
      return;
    }
  } else {
    jobInfo = capitalizeWords(jobInfo);
    usernameCell.setValue(jobInfo);
  }

  var range = activeSheet.getRange(4, binsColumn, activeSheet.getLastRow() - 3, 1);
  var values = range.getValues();

  var columnBValues = homelessGearSheet.getRange("B:B").getValues();
  var existingItems = new Set(columnBValues.flat().filter(String)); // Store existing items for quick lookup

  var barcodes = activeSheet.getRange(4, barcodeColumn, values.length, 1).getValues();
  var itemNames = activeSheet.getRange(4, userColumn, values.length, 1).getValues();

  var itemCounts = {};
  var lostAndFoundItems = [];
  var csvData = [];

  const keywordsRegex = /\b(?:Disposed|Repair|Lost|Inactive|Sale|Pending QC)\b(?=\s*\||$)/i;

  for (var i = 0; i < values.length; i++) {
    var binStatus = values[i][0];
    var itemName = itemNames[i][0];
    var barcode = barcodes[i][0];

    if (
      binStatus === "No Bin" &&
      itemName !== "Item Not Found" &&
      !/case|pelican/i.test(itemName) // Skip items with "case" or "pelican"
    ) {
      if (!existingItems.has(itemName)) {
        if (!itemCounts[itemName]) {
          itemCounts[itemName] = { barcode: barcode, count: 0 };
        }
        itemCounts[itemName].count += 1;
      }
    } else if (["LOST", "DISPOSED", "INACTIVE"].includes(binStatus)) {
      var statusText = binStatus.charAt(0) + binStatus.slice(1).toLowerCase();

      // Remove unwanted keywords in one pass
      itemName = itemName.replace(keywordsRegex, "").replace(/\s+\|/, "|").trim();

      var consigner = "";
      if (itemName.includes("|")) {
        var parts = itemName.split("|");
        itemName = parts[0].trim(); // Keep the item name before the pipe
        consigner = parts[1].trim(); // Store the consigner
      } else {
        itemName = itemName.trim();
        consigner = ""; // still empty initially
      }

      // Clean up dangling pipes or extra spaces
      itemName = itemName.replace(/\|\s*$/, "").trim();

      // Set default consigner if none specified
      if (!itemName.includes("|") && consigner === "") {
        consigner = "Keslow";
      }

      // Clean up dangling pipes or extra spaces again
      itemName = itemName.replace(/\|\s*$/, "").trim();

      // Push to Lost & Found with consigner in column H if barcode doesn't exist
      if (!existingLFBarcodes.has(barcode)) {
        lostAndFoundItems.push([barcode, itemName, statusText, 1, jobInfo, "", consigner]);
      }
    }

    if (barcode) {
      csvData.push([barcode]);
    }
  }

  // Append homeless gear items
  var dataToAppend = [];
  for (var item in itemCounts) {
    var quantity = Math.max(itemCounts[item].count, 1);
    dataToAppend.push([itemCounts[item].barcode, item, "", quantity, jobInfo]);
  }

  if (dataToAppend.length > 0) {
    var lastDataRow = columnBValues.filter(row => row[0].toString().trim() !== "").length;
    var targetRange = homelessGearSheet.getRange(lastDataRow + 1, 1, dataToAppend.length, 5);
    targetRange.setValues(dataToAppend);
  }

  // Append lost and found items
  if (lostAndFoundItems.length > 0) {
    var lastDataRowLF = lostAndFoundSheet.getRange("B:B").getValues().filter(row => row[0].toString().trim() !== "").length;
    var targetRangeLF = lostAndFoundSheet.getRange(lastDataRowLF + 1, 1, lostAndFoundItems.length, 7);
    targetRangeLF.setValues(lostAndFoundItems);

    var checkboxRange = lostAndFoundSheet.getRange(lastDataRowLF + 1, 6, lostAndFoundItems.length, 1);
    checkboxRange.insertCheckboxes();
  }

  // Update analytics
  var statusCount = lostAndFoundItems.length;
  var currentTotal = analyticsSheet.getRange("AA2").getValue() || 0;
  var newTotal = currentTotal + statusCount;
  analyticsSheet.getRange("AA2").setValue(newTotal);
  // Logger.log(`Updated analytics: Added ${statusCount} Lost & Found items, new total is ${newTotal}`); // removed detailed log

  // Count unique barcodes and update Z2
  var uniqueBarcodes = new Set(barcodes.flat().filter(String)).size;
  var currentTotal = analyticsSheet.getRange("Z2").getValue() || 0;
  var newTotal = currentTotal + uniqueBarcodes;
  analyticsSheet.getRange("Z2").setValue(newTotal);
  Logger.log(`Reset Generic Bay complete. Lost & Found added: ${statusCount}, Barcodes processed: ${uniqueBarcodes}.`);

  // Save barcodes to CSV and clear content
  saveBarcodesToCSV(csvData, jobInfo);

  activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1).clearContent();
  usernameCell.clearContent();
}

/**
 * Saves the barcodes to a CSV file in the Shared Drive "Barcode Bay Archives" folder
 * @param {Array} csvData - Array of barcode data to save
 * @param {string} jobInfo - Job info from cell B2 for the filename
 */
function saveBarcodesToCSV(csvData, jobInfo) {
  var folderName = "Barcode Bay Archives";
  var timestamp = new Date().toISOString().replace("T", " ").split(".")[0];
  var fileName = timestamp + " " + jobInfo + ".csv";
  var csvContent = csvData.map(row => row.join(",")).join("\n");

  // Save to Shared Drive
  saveToSharedDrive(folderName, fileName, csvContent);
  
  // Also save to user's personal Drive as backup
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.createFile(fileName, csvContent, MimeType.CSV);
}

/**
 * Helper function to save files to Shared Drive
 * @param {string} folderName - Name of the folder in Shared Drive
 * @param {string} fileName - Name of the file to create
 * @param {string} fileContent - Content of the file
 */
function saveToSharedDrive(folderName, fileName, fileContent) {
  try {
    // TODO: Replace with actual Shared Drive ID
    const SHARED_DRIVE_ID = "YOUR_SHARED_DRIVE_ID_HERE";
    
    // Find or create folder in Shared Drive
    const folders = Drive.Files.list({
      q: `'${SHARED_DRIVE_ID}' in parents and title='${folderName}' and mimeType='application/vnd.google-apps.folder'`
    });
    
    let targetFolderId;
    if (folders.items && folders.items.length > 0) {
      targetFolderId = folders.items[0].id;
    } else {
      // Create folder in Shared Drive
      const newFolder = Drive.Files.insert({
        title: folderName,
        parents: [{id: SHARED_DRIVE_ID}],
        mimeType: 'application/vnd.google-apps.folder'
      });
      targetFolderId = newFolder.id;
    }
    
    // Create file in Shared Drive folder
    Drive.Files.insert({
      title: fileName,
      parents: [{id: targetFolderId}],
      mimeType: 'text/csv'
    }, Utilities.newBlob(fileContent, 'text/csv'));
    
    Logger.log(`✅ File saved to Shared Drive: ${fileName}`);
  } catch (error) {
    Logger.log(`❌ Error saving to Shared Drive: ${error.toString()}`);
    // Continue execution even if Shared Drive save fails
  }
}

function setSelectedUsernameCell(cellA1) {
  PropertiesService.getScriptProperties().setProperty('selectedUsernameCell', cellA1);
} 