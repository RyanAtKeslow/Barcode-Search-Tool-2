/**
 * Reset Receiving Bay - Receiving Bay Reset and Data Processing Script
 * [With Added Logging]
 */

function resetReceivingBay() {
  console.time("ResetReceivingBay_Total_Time"); // Start Timer
  console.log("--- Starting Reset Receiving Bay Script ---");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // Check if we're in the correct sheet
  if (activeSheet.getName() !== "Receiving Bays") {
    console.warn("Script attempted on wrong sheet: " + activeSheet.getName());
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
    console.log("User identified via fetchUserInfoFromEmail: " + userFirstName);
  } catch (error) {
    console.warn("Primary user fetch failed, trying fallback...", error);
    try {
      const { nickname } = fetchUserEmailandNickname();
      userFirstName = nickname;
      console.log("User identified via fallback (nickname): " + userFirstName);
    } catch (error) {
      console.error("Critical Error: Could not identify user.");
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
  
  console.log(`Found ${usernameMatches.length} matching bays for user: ${userFirstName}`);

  // Handle multiple matches
  if (usernameMatches.length > 1) {
    console.log("Multiple matches found, prompting user for selection.");
    const selectedMatch = setSelectedMatch(usernameMatches);
    continueResetShippingBay(selectedMatch.cellA1);
  } else if (usernameMatches.length === 1) {
    console.log("Single match found. Proceeding with bay: " + usernameMatches[0].cellA1);
    continueResetShippingBay(usernameMatches[0].cellA1);
  } else {
    console.warn("No matches found in Row 2 for user: " + userFirstName);
    SpreadsheetApp.getUi().alert("Could not find your name in row 2");
    return;
  }
}

function continueResetShippingBay(selectedCellA1) {
  console.log(`--- Processing Bay at ${selectedCellA1} ---`);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var usernameCell = activeSheet.getRange(selectedCellA1);
  
  var homelessGearSheet = ss.getSheetByName("HOMELESS GEAR");
  var lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  var analyticsSheet = ss.getSheetByName("Analytics");
  
  // Get existing data from Lost & Found
  console.log("Loading existing Lost & Found data...");
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

  // Job Info / Signature Logic
  if (jobInfo === "") {
    console.log("Job Info is empty. Prompting user for signature.");
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "Please Enter Your Name",
      "(Please sign the scanner's name before resetting the bay)",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      console.log("User cancelled signature prompt.");
      ui.alert("Please sign your work to continue");
      return;
    }

    if (response.getResponseText().trim() !== "") {
      jobInfo = capitalizeWords(response.getResponseText().trim());
      usernameCell.setValue(jobInfo);
      console.log("User signed as: " + jobInfo);
    } else {
      ui.alert("You must sign your work to continue");
      return;
    }
  } else {
    jobInfo = capitalizeWords(jobInfo);
    usernameCell.setValue(jobInfo);
    console.log("Job Info present: " + jobInfo);
  }

  // Reading Bay Data
  var range = activeSheet.getRange(4, binsColumn, activeSheet.getLastRow() - 3, 1);
  var values = range.getValues();
  console.log(`Scanning ${values.length} rows in the bay.`);

  var columnBValues = homelessGearSheet.getRange("B:B").getValues();
  var existingItems = new Set(columnBValues.flat().filter(String));

  var barcodes = activeSheet.getRange(4, barcodeColumn, values.length, 1).getValues();
  var itemNames = activeSheet.getRange(4, userColumn, values.length, 1).getValues();

  var itemCounts = {};
  var lostAndFoundItems = [];
  var quantityUpdates = [];
  var csvData = [];
  const keywordsRegex = /\b(?:Disposed|Repair|Lost|Inactive|Sale|Pending QC)\b(?=\s*\||$)/i;

  // --- Main Processing Loop ---
  console.time("Processing_Loop");
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
      }
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
        } else {
          // Pending item logic
          const pendingIndex = lostAndFoundItems.findIndex(item => item[0] === barcode.toString());
          if (pendingIndex !== -1) {
            lostAndFoundItems[pendingIndex][3] += 1;
          }
        }
      } else {
        lostAndFoundItems.push([barcode, itemName, statusText, 1, jobInfo, "", consigner]);
        existingBarcodeMap.set(barcode.toString(), "pending");
      }
    }

    if (barcode) {
      csvData.push([barcode]);
    }
  }
  console.timeEnd("Processing_Loop");

  // --- Appending Data ---

  // Homeless Gear
  var dataToAppend = [];
  for (var item in itemCounts) {
    var quantity = Math.max(itemCounts[item].count, 1);
    dataToAppend.push([itemCounts[item].barcode, item, "", quantity, jobInfo]);
  }

  if (dataToAppend.length > 0) {
    console.log(`Appending ${dataToAppend.length} items to Homeless Gear.`);
    var lastDataRow = columnBValues.filter(row => row[0].toString().trim() !== "").length;
    var targetRange = homelessGearSheet.getRange(lastDataRow + 1, 1, dataToAppend.length, 5);
    targetRange.setValues(dataToAppend);
  } else {
    console.log("No new Homeless Gear items found.");
  }

  // Update quantities for existing L&F items
  if (quantityUpdates.length > 0) {
    console.log(`Updating quantities for ${quantityUpdates.length} existing Lost & Found items.`);
    quantityUpdates.forEach(update => {
      lostAndFoundSheet.getRange(update.row, 4).setValue(update.newQuantity);
      lostAndFoundSheet.getRange(update.row, 5).setValue(update.jobInfo);
    });
  }

  // Append new Lost & Found items
  if (lostAndFoundItems.length > 0) {
    console.log(`Appending ${lostAndFoundItems.length} NEW items to Lost & Found.`);
    var lastDataRowLF = lostAndFoundSheet.getRange("B:B").getValues().filter(row => row[0].toString().trim() !== "").length;
    var targetRangeLF = lostAndFoundSheet.getRange(lastDataRowLF + 1, 1, lostAndFoundItems.length, 7);
    targetRangeLF.setValues(lostAndFoundItems);

    var checkboxRange = lostAndFoundSheet.getRange(lastDataRowLF + 1, 6, lostAndFoundItems.length, 1);
    checkboxRange.insertCheckboxes();
  } else {
    console.log("No new Lost & Found items to append.");
  }

  // --- Update Analytics ---
  var statusCount = lostAndFoundItems.length;
  var currentTotalAA = analyticsSheet.getRange("AA2").getValue() || 0;
  var newTotalAA = currentTotalAA + statusCount;
  analyticsSheet.getRange("AA2").setValue(newTotalAA);
  console.log(`Analytics updated (AA2): Added ${statusCount} L&F items. New Total: ${newTotalAA}`);

  var uniqueBarcodes = new Set(barcodes.flat().filter(String)).size;
  var currentTotalZ = analyticsSheet.getRange("Z2").getValue() || 0;
  var newTotalZ = currentTotalZ + uniqueBarcodes;
  analyticsSheet.getRange("Z2").setValue(newTotalZ);
  console.log(`Analytics updated (Z2): Added ${uniqueBarcodes} unique barcodes. New Total: ${newTotalZ}`);

  // --- CSV & DB Updates ---
  console.log("Initiating CSV Archive...");
  saveBarcodesToCSV(csvData, jobInfo);
  console.log("CSV Archive complete.");

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
  console.log("Clearing Bay content...");
  activeSheet.getRange(4, barcodeColumn, activeSheet.getLastRow() - 3, 1).clearContent();
  usernameCell.clearContent();
  
  console.log("--- Reset Receiving Bay Complete ---");
  console.timeEnd("ResetReceivingBay_Total_Time"); // End Timer
}

/**
 * Saves the barcodes to a CSV file in the Shared Drive "Barcode Bay Archives" folder
 */
function saveBarcodesToCSV(csvData, jobInfo) {
  var folderName = "Barcode Bay Archives";
  var timestamp = new Date().toISOString().replace("T", " ").split(".")[0];
  var fileName = timestamp + " " + jobInfo + ".csv";
  var csvContent = csvData.map(row => row.join(",")).join("\n");

  console.log(`Saving CSV: ${fileName}`);

  // Save to Shared Drive
  saveToSharedDrive(folderName, fileName, csvContent);
  
  // Also save to user's personal Drive as backup
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.createFile(fileName, csvContent, MimeType.CSV);
}

// saveToSharedDrive function is now centralized in Data Model.js

function setSelectedUsernameCell(cellA1) {
  PropertiesService.getScriptProperties().setProperty('selectedUsernameCell', cellA1);
  console.log("Selected username cell stored: " + cellA1);
}