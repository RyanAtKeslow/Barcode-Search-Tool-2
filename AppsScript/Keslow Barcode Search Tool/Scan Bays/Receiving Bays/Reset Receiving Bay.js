/**
 * Reset Receiving Bay - Receiving Bay Reset and Data Processing Script
 * [With Added Logging]
 */

function resetReceivingBay() {
  var startTime = new Date().getTime();
  Logger.log("=== Reset Receiving Bay: Script Started ===");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Receiving Bays") {
    Logger.log("ERROR: Script not run from 'Receiving Bays' sheet. Current sheet: " + activeSheet.getName());
    SpreadsheetApp.getUi().alert("Please run this script from the 'Receiving Bays' sheet");
    return;
  }
  
  Logger.log("Sheet validation passed: Running from 'Receiving Bays' sheet");
  
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
    Logger.log("WARN: Primary user identification failed, trying fallback method. Error: " + error.toString());
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
    continueResetShippingBay(selectedMatch.cellA1, startTime);
  } else if (usernameMatches.length === 1) {
    Logger.log("Single username match found: " + usernameMatches[0].cellA1 + " (" + usernameMatches[0].value + ")");
    continueResetShippingBay(usernameMatches[0].cellA1, startTime);
  } else {
    Logger.log("ERROR: No username matches found in row 2");
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueResetShippingBay(selectedCellA1, startTime) {
  Logger.log("=== continueResetShippingBay: Processing bay reset for cell " + selectedCellA1 + " ===");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");
  
  // Get existing data from Lost & Found
  Logger.log("Loading existing Lost & Found data...");
  const existingData = lostAndFoundSheet.getRange("A:D").getValues();
  const existingBarcodeMap = new Map(); // Map barcode to row index
  
  // Build map of existing barcodes and their row positions
  for (let i = 0; i < existingData.length; i++) {
    const barcode = existingData[i][0];
    if (barcode && barcode.toString().trim() !== "") {
      existingBarcodeMap.set(barcode.toString(), i + 1); // +1 for 1-based row indexing
    }
  }
  
  Logger.log("Loaded " + existingBarcodeMap.size + " existing barcodes from Lost & Found sheet");
  
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

  // Job Info / Signature Logic
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

  // Reading Bay Data
  var range = activeSheet.getRange(4, binsColumn, activeSheet.getLastRow() - 3, 1);
  var values = range.getValues();
  
  Logger.log("Processing " + values.length + " rows of bin status data");

  var columnBValues = homelessGearSheet.getRange("B:B").getValues();
  var existingItems = new Set(columnBValues.flat().filter(String));
  
  Logger.log("Loaded " + existingItems.size + " existing items from HOMELESS GEAR sheet");

  var barcodes = activeSheet.getRange(4, barcodeColumn, values.length, 1).getValues();
  var itemNames = activeSheet.getRange(4, userColumn, values.length, 1).getValues();

  var itemCounts = {};
  var lostAndFoundItems = [];
  var quantityUpdates = [];
  var csvData = [];
  var skippedItems = 0;
  var duplicateLFSkipped = 0;
  const keywordsRegex = /\b(?:Disposed|Repair|Lost|Inactive|Sale|Pending QC)\b(?=\s*\||$)/i;

  // --- Main Processing Loop ---
  Logger.log("Starting item processing loop...");
  var loopStartTime = new Date().getTime();
  
  for (var i = 0; i < values.length; i++) {
    var binStatus = values[i][0];
    var itemName = itemNames[i][0];
    var barcode = barcodes[i][0];

    // Homeless Gear Logic
    if (
      binStatus === "No Bin" &&
      itemName !== "Item Not Found" &&
      !/case|pelican/i.test(itemName)
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
    } 
    // Lost/Disposed Logic
    else if (["LOST", "DISPOSED", "INACTIVE"].includes(binStatus)) {
      var statusText = binStatus.charAt(0) + binStatus.slice(1).toLowerCase();

      itemName = itemName.replace(keywordsRegex, "").replace(/\s+\|/, "|").trim();

      var consigner = "";
      if (itemName.includes("|")) {
        var parts = itemName.split("|");
        itemName = parts[0].trim();
        consigner = parts[1].trim();
      } else {
        itemName = itemName.trim();
        consigner = "";
      }

      itemName = itemName.replace(/\|\s*$/, "").trim();

      if (!itemName.includes("|") && consigner === "") {
        consigner = "Keslow";
      }

      itemName = itemName.replace(/\|\s*$/, "").trim();

      if (existingBarcodeMap.has(barcode.toString())) {
        const rowIndex = existingBarcodeMap.get(barcode.toString());
        if (typeof rowIndex === 'number') {
          const currentQuantity = lostAndFoundSheet.getRange(rowIndex, 4).getValue() || 0;
          quantityUpdates.push({
            row: rowIndex,
            newQuantity: currentQuantity + 1,
            jobInfo: jobInfo
          });
          Logger.log("Updating quantity for existing Lost & Found item at row " + rowIndex + ": " + itemName + " (new qty: " + (currentQuantity + 1) + ")");
        } else {
          // Pending item logic
          const pendingIndex = lostAndFoundItems.findIndex(item => item[0] === barcode.toString());
          if (pendingIndex !== -1) {
            lostAndFoundItems[pendingIndex][3] += 1;
            Logger.log("Incremented quantity for pending Lost & Found item: " + itemName);
          }
        }
      } else {
        lostAndFoundItems.push([barcode, itemName, statusText, 1, jobInfo, "", consigner]);
        existingBarcodeMap.set(barcode.toString(), "pending");
        Logger.log("Added to Lost & Found: " + itemName + " (" + statusText + ") - Consigner: " + consigner);
      }
    }

    if (barcode) {
      csvData.push([barcode]);
    }
  }
  
  var loopEndTime = new Date().getTime();
  var loopDuration = ((loopEndTime - loopStartTime) / 1000).toFixed(2);
  Logger.log("Item processing loop complete. Duration: " + loopDuration + " seconds");
  
  Logger.log("Item processing complete. Summary:");
  Logger.log("  - Homeless gear items: " + Object.keys(itemCounts).length);
  Logger.log("  - Lost & Found items (new): " + lostAndFoundItems.length);
  Logger.log("  - Lost & Found quantity updates: " + quantityUpdates.length);
  Logger.log("  - Skipped items (case/pelican): " + skippedItems);
  Logger.log("  - Total barcodes for CSV: " + csvData.length);

  // --- Appending Data ---

  // Homeless Gear
  var dataToAppend = [];
  for (var item in itemCounts) {
    var quantity = Math.max(itemCounts[item].count, 1);
    dataToAppend.push([itemCounts[item].barcode, item, "", quantity, jobInfo]);
  }

  if (dataToAppend.length > 0) {
    var lastDataRow = columnBValues.filter(row => row[0].toString().trim() !== "").length;
    var targetRange = homelessGearSheet.getRange(lastDataRow + 1, 1, dataToAppend.length, 5);
    targetRange.setValues(dataToAppend);
    Logger.log("Appended " + dataToAppend.length + " items to HOMELESS GEAR sheet starting at row " + (lastDataRow + 1));
  } else {
    Logger.log("No homeless gear items to append");
  }

  // Update quantities for existing L&F items
  if (quantityUpdates.length > 0) {
    Logger.log("Updating quantities for " + quantityUpdates.length + " existing Lost & Found items");
    quantityUpdates.forEach(update => {
      lostAndFoundSheet.getRange(update.row, 4).setValue(update.newQuantity);
      lostAndFoundSheet.getRange(update.row, 5).setValue(update.jobInfo);
    });
    Logger.log("Completed quantity updates for " + quantityUpdates.length + " items");
  } else {
    Logger.log("No Lost & Found quantity updates needed");
  }

  // Append new Lost & Found items
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

  // --- Update Analytics ---
  var statusCount = lostAndFoundItems.length;
  var currentTotalAA = analyticsSheet.getRange("AA2").getValue() || 0;
  var newTotalAA = currentTotalAA + statusCount;
  analyticsSheet.getRange("AA2").setValue(newTotalAA);
  Logger.log("Updated Lost & Found analytics (AA2): Added " + statusCount + " items, new total is " + newTotalAA);

  var uniqueBarcodes = new Set(barcodes.flat().filter(String)).size;
  var currentTotalZ = analyticsSheet.getRange("Z2").getValue() || 0;
  var newTotalZ = currentTotalZ + uniqueBarcodes;
  analyticsSheet.getRange("Z2").setValue(newTotalZ);
  Logger.log("Updated barcode analytics (Z2): Added " + uniqueBarcodes + " unique barcodes, new total is " + newTotalZ);

  // --- CSV & DB Updates ---
  Logger.log("Saving barcodes to CSV archive...");
  saveBarcodesToCSV(csvData, jobInfo);

  // ============================================================================
  // COMMENTED OUT: Camera Status "Returned" Update Functionality
  // This code previously updated camera status to "Returned" when barcodes were
  // scanned in the receiving bay. Commented out for potential future reference.
  // To re-enable: uncomment the lines below (remove the // from each line)
  // ============================================================================
  // const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
  // 
  // const allBarcodes = Array.from(new Set(barcodes.map(row => row[0]).filter(barcode => barcode && !barcode.toString().includes("Above was exported @"))));
  // 
  // // Re-enabled logging for SendStatus preparation
  // console.log('=== Preparing SendStatus ===');
  // console.log('Barcodes to process: ' + allBarcodes.length);
  // console.log('Username: ' + username);
  // console.log('Job Info: ' + jobInfo);
  // 
  // SendStatus("Returned", allBarcodes, username, jobInfo, userEmail);
  // console.log("SendStatus function called.");
  // ============================================================================

  // --- Cleanup ---
  Logger.log("Clearing bay data from columns...");
  activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1).clearContent();
  usernameCell.clearContent();
  
  var endTime = new Date().getTime();
  var totalDuration = ((endTime - startTime) / 1000).toFixed(2);
  Logger.log("=== Reset Receiving Bay: Script Complete ===");
  Logger.log("Summary - Lost & Found added: " + statusCount + ", Quantity updates: " + quantityUpdates.length + ", Barcodes processed: " + uniqueBarcodes + ", Homeless gear added: " + dataToAppend.length);
  Logger.log("Total execution time: " + totalDuration + " seconds");
}

/**
 * Saves the barcodes to a CSV file in the Shared Drive "Barcode Bay Archives" folder
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
  Logger.log("Selected username cell stored: " + cellA1);
}