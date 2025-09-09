/**
 * Script to reset a prep bay, handling homeless gear, lost & found items, and analytics.
 * This script processes items based on their bin status and maintains records in various sheets.
 */

function resetShippingBay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Receiving Bays") {
    SpreadsheetApp.getUi().alert("Please run this script from the 'Receiving Bays' sheet");
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
    continueResetShippingBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    continueResetShippingBay(usernameMatches[0].cellA1);
  } else {
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueResetShippingBay(selectedCellA1) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");
  
  // Get existing data from Lost & Found to check for duplicates and update quantities
  const existingData = lostAndFoundSheet.getRange("A:D").getValues();
  const existingBarcodeMap = new Map(); // Map barcode to row index
  
  // Build map of existing barcodes and their row positions
  for (let i = 0; i < existingData.length; i++) {
    const barcode = existingData[i][0];
    if (barcode && barcode.toString().trim() !== "") {
      existingBarcodeMap.set(barcode.toString(), i + 1); // +1 for 1-based row indexing
    }
  }
  
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
  var quantityUpdates = []; // Track quantity updates for existing items
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

      // Remove unwanted keywords in one pass (loop eliminated)
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

      // Check if barcode already exists in Lost & Found
      if (existingBarcodeMap.has(barcode.toString())) {
        const rowIndex = existingBarcodeMap.get(barcode.toString());
        // Only process if it's a valid row number (not "pending")
        if (typeof rowIndex === 'number') {
          // Increment quantity for existing item
          const currentQuantity = lostAndFoundSheet.getRange(rowIndex, 4).getValue() || 0;
          quantityUpdates.push({
            row: rowIndex,
            newQuantity: currentQuantity + 1,
            jobInfo: jobInfo
          });
        } else {
          // This is a "pending" item from this batch, increment its quantity in the pending items
          const pendingIndex = lostAndFoundItems.findIndex(item => item[0] === barcode.toString());
          if (pendingIndex !== -1) {
            lostAndFoundItems[pendingIndex][3] += 1; // Increment quantity (column D)
          }
        }
      } else {
        // Add new item to Lost & Found
        lostAndFoundItems.push([barcode, itemName, statusText, 1, jobInfo, "", consigner]);
        // Add to map to track for subsequent duplicates in this batch
        existingBarcodeMap.set(barcode.toString(), "pending");
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

  // Update quantities for existing items
  quantityUpdates.forEach(update => {
    lostAndFoundSheet.getRange(update.row, 4).setValue(update.newQuantity);
    lostAndFoundSheet.getRange(update.row, 5).setValue(update.jobInfo); // Overwrite column E
  });

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
  // Suppress verbose logs in analytics section
  // Logger.log(`Updated analytics: Added ${statusCount} Lost & Found items, new total is ${newTotal}`);

  // Count unique barcodes and update Z2
  var uniqueBarcodes = new Set(barcodes.flat().filter(String)).size;
  var currentTotal = analyticsSheet.getRange("Z2").getValue() || 0;
  var newTotal = currentTotal + uniqueBarcodes;
  analyticsSheet.getRange("Z2").setValue(newTotal);
  Logger.log(`Reset Shipping Bay complete. Lost & Found added: ${statusCount}, Barcodes processed: ${uniqueBarcodes}.`);

  // Save barcodes to CSV and clear content
  saveBarcodesToCSV(csvData, jobInfo);

  // Get username from email (for the SendStatus function)
  const { email: userEmail, nickname: username } = fetchUserEmailandNickname();

  // Send status "Returned" to incoming data sheet
  // Pass all barcodes, username, and job info
  const allBarcodes = Array.from(new Set(barcodes.map(row => row[0]).filter(barcode => barcode && !barcode.toString().includes("Above was exported @"))));
  // Comment out detailed SendStatus debug logs
  // Logger.log('=== About to call SendStatus ===');
  // Logger.log('Current Script ID: ' + ScriptApp.getScriptId());
  // Logger.log('Number of barcodes to process: ' + allBarcodes.length);
  // Logger.log('Username: ' + username);
  // Logger.log('Job Info: ' + jobInfo);
  SendStatus("Returned", allBarcodes, username, jobInfo, userEmail);

  activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1).clearContent();
  usernameCell.clearContent();
}

/**
 * Saves the barcodes to a CSV file in the "Barcode Bay Archives" folder
 * @param {Array} csvData - Array of barcode data to save
 * @param {string} jobInfo - Job info from cell B2 for the filename
 */
function saveBarcodesToCSV(csvData, jobInfo) {
  var folderName = "Barcode Bay Archives";
  var timestamp = new Date().toISOString().replace("T", " ").split(".")[0];
  var fileName = timestamp + " " + jobInfo + ".csv";

  var csvContent = csvData.map(row => row.join(",")).join("\n");

  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  folder.createFile(fileName, csvContent, MimeType.CSV);
}

function setSelectedUsernameCell(cellA1) {
  PropertiesService.getScriptProperties().setProperty('selectedUsernameCell', cellA1);
} 