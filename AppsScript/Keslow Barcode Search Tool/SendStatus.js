// Helper function to normalize barcodes by removing pipes and trimming
function normalizeBarcode(barcode) {
  if (!barcode) return '';
  return barcode.toString().trim();
}

// Helper function to split pipe-delimited barcodes and normalize each one
function splitAndNormalizeBarcodes(barcodeString) {
  if (!barcodeString) return [];
  return barcodeString.toString()
    .split('|')
    .map(b => normalizeBarcode(b))
    .filter(b => b.length > 0);  // Remove empty strings
}

// Helper function to extract job name from username cell
function extractJobName(usernameCellValue, username) {
  if (!usernameCellValue) return '';
  // Remove the username from the cell value and trim
  return usernameCellValue.toString()
    .replace(username, '')
    .trim();
}

// Column definitions for Incoming Database Data sheet
const IncomingCOLS = {
  TIMESTAMP: 0,    // Column A
  BARCODE: 1,      // Column B
  EQUIP_NAME: 2,   // Column C
  STATUS: 3,       // Column D
  USERNAME: 4,     // Column E
  JOB_NAME: 5,     // Column F
  LOCATION: 6      // Column G
};

// Valid status values
const VALID_STATUSES = {
  SHIPPED: "Shipped",
  RETURNED: "Returned",
  PULLED: "Pulled"
};

/**
 * Fetches the user's location from the UserLocationUtilityUS or UserLocationUtilityCAN sheet in the external workbook
 * @param {string} userEmail - The user's email address
 * @param {SpreadsheetApp.Spreadsheet} externalWorkbook - The external workbook object
 * @returns {string} The user's location or 'Unknown' if not found
 */
function getUserLocationByEmail(userEmail, externalWorkbook) {
  const sheetsToCheck = ["UserLocationUtilityUS", "UserLocationUtilityCAN"];
  for (const sheetName of sheetsToCheck) {
    try {
      const utilitySheet = externalWorkbook.getSheetByName(sheetName);
      if (!utilitySheet) {
        Logger.log(`❌ ${sheetName} sheet not found in external workbook.`);
        continue;
      }
      const data = utilitySheet.getRange(2, 1, utilitySheet.getLastRow() - 1, 5).getValues();
      for (let i = 0; i < data.length; i++) {
        // Email is in column E (index 4), location in column A (index 0)
        if (data[i][4] && data[i][4].toString().toLowerCase() === userEmail.toLowerCase()) {
          Logger.log(`✅ Found location for ${userEmail} in ${sheetName}: ${data[i][0]}`);
          return data[i][0] || "Unknown";
        }
      }
    } catch (err) {
      Logger.log(`❌ Error fetching user location from ${sheetName}: ${err}`);
    }
  }
  Logger.log(`User's location could not be derived from their email: ${userEmail}`);
  return "Unknown";
}

/**
 * Core function to process digital cameras and send their status to the database
 * @param {string} status - The status to set (must be one of VALID_STATUSES)
 * @param {Array<string>} barcodes - Array of barcodes to check
 * @param {string} username - Username for the database entry
 * @param {string} jobName - Job name for the database entry
 * @param {string} userEmail - User's email for location lookup
 * @returns {void}
 */
function SendStatus(status, barcodes, username, jobName, userEmail) {
  Logger.log('🔵🔵🔵 NEW VERSION OF SENDSTATUS RUNNING 🔵🔵🔵');
  Logger.log('=== SendStatus VERSION 2.0 START ===');
  Logger.log('Script ID: ' + ScriptApp.getScriptId());
  Logger.log('Status: ' + status);
  Logger.log('Number of barcodes received: ' + barcodes.length);
  Logger.log('Username: ' + username);
  Logger.log('Job Name: ' + jobName);
  
  if (!Object.values(VALID_STATUSES).includes(status)) {
    throw new Error(`Invalid status: ${status}. Must be one of: ${Object.values(VALID_STATUSES).join(", ")}`);
  }

  try {
    Logger.log(`=== SendStatus START (Status: ${status}) ===`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const barcodeDictSheet = ss.getSheetByName("barcode dictionary");
    
    // Open the external workbook for Cameras sheet
    const externalWorkbook = SpreadsheetApp.openById("13PMB5l5PJr4HHQ0W9A7KvTu2derCVtLKRHtoJ-2LxW4");
    Logger.log('Writing to external spreadsheet ID: 13PMB5l5PJr4HHQ0W9A7KvTu2derCVtLKRHtoJ-2LxW4');
    const incomingDataSheet = externalWorkbook.getSheetByName("Cameras");

    // Check each sheet individually and provide specific error messages
    if (!barcodeDictSheet) {
      throw new Error("Required sheet 'barcode dictionary' not found in the active spreadsheet");
    }
    if (!incomingDataSheet) {
      throw new Error("Required sheet 'Cameras' not found in the external workbook");
    }

    Logger.log('✅ Successfully connected to external workbook');
    Logger.log(`Processing for username: ${username}`);
    Logger.log(`Job name: ${jobName}`);

    // Get barcode dictionary data
    const dictData = barcodeDictSheet.getDataRange().getValues();
    Logger.log('=== Scanning barcode dictionary ===');
    Logger.log(`Total rows in dictionary: ${dictData.length}`);
    
    let digitalCameras = [];
    
    // Log all barcodes to be checked
    Logger.log('=== Barcodes to check ===');
    barcodes.forEach(barcode => {
      if (!barcode) return;
      const normalizedBarcode = normalizeBarcode(barcode);
      Logger.log(`Found barcode: "${barcode}" (type: ${typeof barcode}, length: ${barcode.toString().length})`);
      Logger.log(`Normalized to: "${normalizedBarcode}" (type: ${typeof normalizedBarcode}, length: ${normalizedBarcode.length})`);
    });
    Logger.log(`Total barcodes to check: ${barcodes.length}`);

    barcodes.forEach(barcode => {
      if (!barcode) return;
      const normalizedBarcode = normalizeBarcode(barcode);
      Logger.log(`\nChecking barcode: "${barcode}"`);
      Logger.log(`Normalized to: "${normalizedBarcode}"`);
      let found = false;
      let matchAttempts = 0;
      
      for (let i = 1; i < dictData.length; i++) { // skip header
        const category = dictData[i][0];          // Column A - Category
        const equipName = dictData[i][3];         // Column D - Equipment Name
        const rawDictBarcodes = dictData[i][6];   // Column G - Raw barcodes string
        
        // Split the pipe-delimited barcodes and check each one
        const dictBarcodes = splitAndNormalizeBarcodes(rawDictBarcodes);
        
        // Only log the first 5 comparison attempts to avoid flooding the logs
        if (matchAttempts < 5) {
          Logger.log(`\nChecking dictionary row ${i + 1}:`);
          Logger.log(`Raw dictionary barcodes: "${rawDictBarcodes}"`);
          Logger.log(`Split into: ${JSON.stringify(dictBarcodes)}`);
          Logger.log(`Looking for: "${normalizedBarcode}"`);
          matchAttempts++;
        }
        
        // Check if our barcode matches any of the split barcodes
        if (dictBarcodes.includes(normalizedBarcode)) {
          Logger.log(`\n✅ MATCH FOUND in row ${i + 1}:`);
          Logger.log(`- Category: ${category}`);
          Logger.log(`- Equipment Name: ${equipName}`);
          Logger.log(`- Matching barcode: "${normalizedBarcode}"`);
          Logger.log(`- From barcode list: "${rawDictBarcodes}"`);
          
          if (category == "Digital Cameras") {
            digitalCameras.push({
              barcode: barcode,
              equipmentName: equipName
            });
            Logger.log('✅ Added to digital cameras list');
            found = true;
            break;
          } else {
            Logger.log('❌ Not a digital camera, skipping');
          }
        }
      }
      
      if (!found) {
        Logger.log(`❌ No dictionary match found for barcode: "${barcode}"`);
        Logger.log(`Attempted ${matchAttempts} comparisons (showing first 5 in logs)`);
      }
    });

    if (digitalCameras.length === 0) {
      Logger.log('No digital cameras found in scanned barcodes');
      return;
    }

    // Append data to Cameras sheet in external workbook
    Logger.log('Appending data to external Cameras sheet...');
    const timestamp = new Date().toLocaleString();
    digitalCameras.forEach(camera => {
      const newRow = Array(7).fill(''); // Initialize array with 7 empty strings
      newRow[IncomingCOLS.TIMESTAMP] = timestamp;
      newRow[IncomingCOLS.BARCODE] = camera.barcode;
      newRow[IncomingCOLS.EQUIP_NAME] = camera.equipmentName;
      newRow[IncomingCOLS.STATUS] = status;
      newRow[IncomingCOLS.USERNAME] = username;
      newRow[IncomingCOLS.JOB_NAME] = jobName;
      newRow[IncomingCOLS.LOCATION] = getUserLocationByEmail(userEmail, externalWorkbook);
      try {
        incomingDataSheet.appendRow(newRow);
        Logger.log(`✅ Appended row to external sheet:`);
        Logger.log(`- Timestamp: ${newRow[IncomingCOLS.TIMESTAMP]}`);
        Logger.log(`- Barcode: ${newRow[IncomingCOLS.BARCODE]}`);
        Logger.log(`- Equipment Name: ${newRow[IncomingCOLS.EQUIP_NAME]}`);
        Logger.log(`- Status: ${newRow[IncomingCOLS.STATUS]}`);
        Logger.log(`- Username: ${newRow[IncomingCOLS.USERNAME]}`);
        Logger.log(`- Job Name: ${newRow[IncomingCOLS.JOB_NAME]}`);
        Logger.log(`- Location: ${newRow[IncomingCOLS.LOCATION]}`);
      } catch (appendError) {
        Logger.log(`❌ Error appending row: ${appendError.toString()}`);
        throw appendError;
      }
    });

    Logger.log(`✅ Successfully processed ${digitalCameras.length} digital cameras`);
    Logger.log('=== SendStatus END ===');
    
  } catch (error) {
    Logger.log(`❌ Error in SendStatus: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Sends digital cameras with "Shipped" status
 */
function ShippedStatusSending() {
  SendStatus(VALID_STATUSES.SHIPPED);
}

/**
 * Sends digital cameras with "Returned" status
 */
function ReturnedStatusSending() {
  SendStatus(VALID_STATUSES.RETURNED);
}

/**
 * Sends digital cameras with "Pulled" status
 */
function PulledStatusSending() {
  SendStatus(VALID_STATUSES.PULLED);
} 