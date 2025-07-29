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
 * Returns a lookup map { barcode -> {category, equipName} } for Digital Cameras.
 * Cached in CacheService for 6 hours using the sheet's size as the version key.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dictSheet
 * @returns {Object}
 */
function getBarcodeLookup(dictSheet) {
  const cache = CacheService.getScriptCache();
  const versionKey = `dict_${dictSheet.getLastRow()}_${dictSheet.getLastColumn()}`;
  const cached = cache.get(versionKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const data = dictSheet.getDataRange().getValues();
  const lookup = {};
  for (let i = 1; i < data.length; i++) { // skip header row
    const category = data[i][0];
    if (category !== 'Digital Cameras') continue; // only store needed rows
    const equipName = data[i][3];
    const rawBarcodes = data[i][6];
    splitAndNormalizeBarcodes(rawBarcodes).forEach(bc => {
      if (!lookup[bc]) {
        lookup[bc] = { category, equipName };
      }
    });
  }

  // store for 6h (21600 s) – limited by 100 KB cache size
  try { cache.put(versionKey, JSON.stringify(lookup), 21600); } catch (e) {}
  return lookup;
}

/**
 * Builds / retrieves a map { email -> location } from the two utility sheets.
 * Cached for 6 hours to avoid repeated scans of the external workbook.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} externalWb
 * @returns {Object}
 */
function getUserLocationMap(externalWb) {
  const cache = CacheService.getScriptCache();
  const key = `userLoc_${externalWb.getId()}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const map = {};
  ['UserLocationUtilityUS', 'UserLocationUtilityCAN'].forEach(sheetName => {
    const sh = externalWb.getSheetByName(sheetName);
    if (!sh) return;
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    vals.forEach(r => {
      const email = (r[4] || '').toString().toLowerCase();
      if (email) map[email] = r[0] || 'Unknown';
    });
  });
  try { cache.put(key, JSON.stringify(map), 21600); } catch (e) {}
  return map;
}

/**
 * Fetches the user's location from the UserLocationUtilityUS or UserLocationUtilityCAN sheet in the external workbook
 * @param {string} userEmail - The user's email address
 * @param {SpreadsheetApp.Spreadsheet} externalWorkbook - The external workbook object
 * @returns {string} The user's location or 'Unknown' if not found
 */
function getUserLocationByEmail(userEmail, externalWorkbook) {
  const map = getUserLocationMap(externalWorkbook);
  return map[userEmail.toLowerCase()] || 'Unknown';
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
  Logger.log(`SendStatus start | status: ${status} | barcodes: ${barcodes.length} | user: ${username}`);
  
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

    // Retrieve (or build) cached barcode lookup map
    const barcodeLookup = getBarcodeLookup(barcodeDictSheet);
    const digitalCameras = [];
    barcodes.forEach(barcode => {
      if (!barcode) return;
      const normalized = normalizeBarcode(barcode);
      const match = barcodeLookup[normalized];
      if (match && match.category === 'Digital Cameras') {
        digitalCameras.push({ barcode, equipmentName: match.equipName || match.equipmentName || match.equipName });
      }
    });

    if (digitalCameras.length === 0) {
      Logger.log('No digital cameras found in scanned barcodes');
      return;
    }

    // Append data to Cameras sheet in external workbook
    const timestamp = new Date().toLocaleString();
    const rowsToAppend = digitalCameras.map(cam => {
      const row = Array(7).fill('');
      row[IncomingCOLS.TIMESTAMP] = timestamp;
      row[IncomingCOLS.BARCODE] = cam.barcode;
      row[IncomingCOLS.EQUIP_NAME] = cam.equipmentName;
      row[IncomingCOLS.STATUS] = status;
      row[IncomingCOLS.USERNAME] = username;
      row[IncomingCOLS.JOB_NAME] = jobName;
      row[IncomingCOLS.LOCATION] = getUserLocationByEmail(userEmail, externalWorkbook);
      return row;
    });

    if (rowsToAppend.length > 0) {
      const startRow = incomingDataSheet.getLastRow() + 1;
      incomingDataSheet.getRange(startRow, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
    }
    Logger.log(`SendStatus complete | digital cameras appended: ${rowsToAppend.length}`);
    
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