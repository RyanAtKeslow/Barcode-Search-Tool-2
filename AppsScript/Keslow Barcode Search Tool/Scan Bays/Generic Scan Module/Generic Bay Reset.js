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
  Logger.log("=== Reset Generic Bay: Script Started ===");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Barcode SEARCH") {
    Logger.log("ERROR: Script not run from 'Barcode SEARCH' sheet. Current sheet: " + activeSheet.getName());
    SpreadsheetApp.getUi().alert("Please run this script from the 'Barcode SEARCH' sheet");
    return;
  }
  
  Logger.log("Sheet validation passed: Running from 'Barcode SEARCH' sheet");
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");

  // Try to get user info with fallback logic
  let userFirstName;
  try {
    const userInfo = fetchUserInfoFromEmail();
    userFirstName = userInfo.firstName;
    Logger.log("User identified via fetchUserInfoFromEmail: " + userFirstName);
  } catch (error) {
    Logger.log("Primary user identification failed, trying fallback method. Error: " + error.toString());
    try {
      const { nickname } = fetchUserEmailandNickname();
      userFirstName = nickname;
      Logger.log("User identified via fetchUserEmailandNickname: " + userFirstName);
    } catch (error) {
      Logger.log("ERROR: Could not identify user. Both identification methods failed.");
      SpreadsheetApp.getUi().alert("Could not identify user. Please ensure you are logged in with your company email.");
      return;
    }
  }
  
  // Find all username matches in row 2
  var row2Range = activeSheet.getRange(2, 1, 1, activeSheet.getLastColumn());
  var row2Values = row2Range.getValues()[0];
  var usernameMatches = [];
  
  Logger.log("Searching for username matches in row 2. Looking for: " + userFirstName);
  
  for (let j = 0; j < row2Values.length; j++) {
    if (row2Values[j] && row2Values[j].toString().toLowerCase().includes(userFirstName.toLowerCase())) {
      usernameMatches.push({
        cellA1: activeSheet.getRange(2, j + 1).getA1Notation(),
        value: row2Values[j]
      });
    }
  }
  
  Logger.log("Found " + usernameMatches.length + " username match(es) in row 2");
  
  // Handle multiple matches
  if (usernameMatches.length > 1) {
    Logger.log("Multiple username matches found. Prompting user for selection.");
    const selectedMatch = setSelectedMatch(usernameMatches);
    Logger.log("User selected: " + selectedMatch.cellA1);
    continueResetGenericBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    Logger.log("Single username match found: " + usernameMatches[0].cellA1 + " (" + usernameMatches[0].value + ")");
    continueResetGenericBay(usernameMatches[0].cellA1);
  } else {
    Logger.log("ERROR: No username matches found in row 2");
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueResetGenericBay(selectedCellA1) {
  Logger.log("=== continueResetGenericBay: Processing bay reset for cell " + selectedCellA1 + " ===");
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
  
  Logger.log("Loaded " + existingLFBarcodes.size + " existing barcodes from Lost & Found sheet");
  
  var jobInfo = usernameCell.getValue().toString().trim();

  var userColumn = usernameCell.getColumn();
  var barcodeColumn = userColumn - 1;
  var binsColumn = userColumn + 1;
  
  Logger.log("Column positions - User: " + userColumn + ", Barcode: " + barcodeColumn + ", Bins: " + binsColumn);

  function capitalizeWords(name) {
    return name.replace(/\b\w/g, function (char) {
      return char.toUpperCase();
    });
  }

  if (jobInfo === "") {
    Logger.log("Job info is empty. Prompting user for name.");
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Please Enter Your Name",
      "(Please sign the scanner's name before resetting the bay)",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log("User cancelled name prompt");
      ui.alert("Please sign your work to continue");
      return;
    }

    if (response.getResponseText().trim() !== "") {
      jobInfo = capitalizeWords(response.getResponseText().trim());
      usernameCell.setValue(jobInfo);
      Logger.log("User entered name: " + jobInfo);
    } else {
      Logger.log("User entered empty name. Aborting.");
      ui.alert("You must sign your work to continue");
      return;
    }
  } else {
    jobInfo = capitalizeWords(jobInfo);
    usernameCell.setValue(jobInfo);
    Logger.log("Using existing job info: " + jobInfo);
  }

  var range = activeSheet.getRange(4, binsColumn, activeSheet.getLastRow() - 3, 1);
  var values = range.getValues();
  
  Logger.log("Processing " + values.length + " rows of bin status data");

  var columnBValues = homelessGearSheet.getRange("B:B").getValues();
  var existingItems = new Set(columnBValues.flat().filter(String)); // Store existing items for quick lookup
  
  Logger.log("Loaded " + existingItems.size + " existing items from HOMELESS GEAR sheet");

  var barcodes = activeSheet.getRange(4, barcodeColumn, values.length, 1).getValues();
  var itemNames = activeSheet.getRange(4, userColumn, values.length, 1).getValues();

  var itemCounts = {};
  var lostAndFoundItems = [];
  var csvData = [];
  var skippedItems = 0;
  var duplicateLFSkipped = 0;

  const keywordsRegex = /\b(?:Disposed|Repair|Lost|Inactive|Sale|Pending QC)\b(?=\s*\||$)/i;

  Logger.log("Starting item processing loop...");
  
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
      } else {
        Logger.log("Skipping existing item in HOMELESS GEAR: " + itemName);
      }
    } else if (/case|pelican/i.test(itemName)) {
      skippedItems++;
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
        Logger.log("Added to Lost & Found: " + itemName + " (" + statusText + ") - Consigner: " + consigner);
      } else {
        duplicateLFSkipped++;
        Logger.log("Skipping duplicate barcode in Lost & Found: " + barcode + " (" + itemName + ")");
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

  Logger.log("Item processing complete. Summary:");
  Logger.log("  - Homeless gear items: " + dataToAppend.length);
  Logger.log("  - Lost & Found items: " + lostAndFoundItems.length);
  Logger.log("  - Skipped items (case/pelican): " + skippedItems);
  Logger.log("  - Duplicate Lost & Found skipped: " + duplicateLFSkipped);
  Logger.log("  - Total barcodes for CSV: " + csvData.length);

  if (dataToAppend.length > 0) {
    var lastDataRow = columnBValues.filter(row => row[0].toString().trim() !== "").length;
    var targetRange = homelessGearSheet.getRange(lastDataRow + 1, 1, dataToAppend.length, 5);
    targetRange.setValues(dataToAppend);
    Logger.log("Appended " + dataToAppend.length + " items to HOMELESS GEAR sheet starting at row " + (lastDataRow + 1));
  } else {
    Logger.log("No homeless gear items to append");
  }

  // Append lost and found items
  if (lostAndFoundItems.length > 0) {
    var lastDataRowLF = lostAndFoundSheet.getRange("B:B").getValues().filter(row => row[0].toString().trim() !== "").length;
    var targetRangeLF = lostAndFoundSheet.getRange(lastDataRowLF + 1, 1, lostAndFoundItems.length, 7);
    targetRangeLF.setValues(lostAndFoundItems);

    var checkboxRange = lostAndFoundSheet.getRange(lastDataRowLF + 1, 6, lostAndFoundItems.length, 1);
    checkboxRange.insertCheckboxes();
    Logger.log("Appended " + lostAndFoundItems.length + " items to Lost & Found sheet starting at row " + (lastDataRowLF + 1));
  } else {
    Logger.log("No Lost & Found items to append");
  }

  // Update analytics
  var statusCount = lostAndFoundItems.length;
  var currentTotalLF = analyticsSheet.getRange("AA2").getValue() || 0;
  var newTotalLF = currentTotalLF + statusCount;
  analyticsSheet.getRange("AA2").setValue(newTotalLF);
  Logger.log("Updated Lost & Found analytics (AA2): Added " + statusCount + " items, new total is " + newTotalLF);

  // Count unique barcodes and update Z2
  var uniqueBarcodes = new Set(barcodes.flat().filter(String)).size;
  var currentTotalBarcodes = analyticsSheet.getRange("Z2").getValue() || 0;
  var newTotalBarcodes = currentTotalBarcodes + uniqueBarcodes;
  analyticsSheet.getRange("Z2").setValue(newTotalBarcodes);
  Logger.log("Updated barcode analytics (Z2): Added " + uniqueBarcodes + " unique barcodes, new total is " + newTotalBarcodes);

  // Save barcodes to CSV and clear content
  Logger.log("Saving barcodes to CSV archive...");
  saveBarcodesToCSV(csvData, jobInfo);

  Logger.log("Clearing bay data from columns...");
  activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1).clearContent();
  usernameCell.clearContent();
  
  Logger.log("=== Reset Generic Bay: Script Complete ===");
  Logger.log("Summary - Lost & Found added: " + statusCount + ", Barcodes processed: " + uniqueBarcodes + ", Homeless gear added: " + dataToAppend.length);
}

/**
 * Saves the barcodes to a CSV file in the Shared Drive "Barcode Bay Archives" folder
 * @param {Array} csvData - Array of barcode data to save
 * @param {string} jobInfo - Job info from cell B2 for the filename
 */
function saveBarcodesToCSV(csvData, jobInfo) {
  Logger.log("saveBarcodesToCSV: Starting CSV save process");
  var folderName = "Barcode Bay Archives";
  var timestamp = getConsistentTimestamp();
  var fileName = timestamp + "_" + jobInfo + ".csv";
  var csvContent = csvData.map(row => row.join(",")).join("\n");
  
  Logger.log("CSV file details - Name: " + fileName + ", Rows: " + csvData.length);

  // Save to Shared Drive
  try {
    saveToSharedDrive(folderName, fileName, csvContent);
    Logger.log("Successfully saved CSV to Shared Drive: " + folderName);
  } catch (error) {
    Logger.log("ERROR saving to Shared Drive: " + error.toString());
  }
  
  // Also save to user's personal Drive as backup
  try {
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    folder.createFile(fileName, csvContent, MimeType.CSV);
    Logger.log("Successfully saved CSV backup to personal Drive: " + folderName);
  } catch (error) {
    Logger.log("ERROR saving to personal Drive: " + error.toString());
  }
}

// saveToSharedDrive function is now centralized in Data Model.js

function setSelectedUsernameCell(cellA1) {
  PropertiesService.getScriptProperties().setProperty('selectedUsernameCell', cellA1);
} 