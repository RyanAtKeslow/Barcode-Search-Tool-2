/**
 * F2 Import and Process
 * 
 * This script:
 * 1. Scans Google Drive folder for new F2 Service Board Excel exports
 * 2. Converts Excel files to Google Sheets
 * 3. Filters by equipment category (Digital Cameras, 35mm Cameras, 16mm Cameras)
 * 4. Looks up Serial Numbers from Barcode & Serial Database
 * 5. Cross-references with Equipment Scheduling Chart for verification
 * 6. Cross-references with Prep Bay Assignments (secondary check)
 * 7. Writes to "F2 Imports" sheet, updating existing records by barcode
 * 8. Creates alerts for mismatches
 * 9. Moves processed files to "Processed" subfolder
 * 10. Updates RTR Database with Order #s
 * 
 * File naming pattern: "Service [date] at [time].xlsx"
 * Example: "Service 2025-12-03 at 3.41.39 PM.xlsx"
 */

// Configuration
const F2_IMPORT_FOLDER_ID = '1nUy7lWNr1BVCAxyLsnFASCTszQpjgEnd';
const FILE_NAME_PATTERN = /^Service \d{4}-\d{2}-\d{2} at \d{1,2}\.\d{2}\.\d{2} (AM|PM)\.xlsx$/i;

// Spreadsheet IDs
const F2_DESTINATION_SPREADSHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
const EQUIPMENT_SCHEDULING_CHART_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';
const PREP_BAY_SPREADSHEET_ID = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';

// Equipment categories to process
const VALID_EQUIPMENT_CATEGORIES = ['Digital Cameras', '35mm Cameras', '16mm Cameras'];

// Sheet names
const F2_IMPORTS_SHEET_NAME = 'F2 Imports';
const BARCODE_SERIAL_SHEET_NAME = 'Barcode & Serial Database';

/**
 * Main function to process F2 imports
 * Can be triggered manually or via time-driven trigger
 */
function processF2Imports() {
  Logger.log("🚀 Starting F2 Import and Process");
  
  try {
    // Step 1: Get the F2 Import folder
    Logger.log("📁 Accessing F2 Import folder...");
    const folder = DriveApp.getFolderById(F2_IMPORT_FOLDER_ID);
    Logger.log(`✅ Folder accessed: ${folder.getName()}`);
    
    // Step 2: Find unprocessed Excel files
    Logger.log("🔍 Searching for unprocessed Excel files...");
    const unprocessedFiles = findUnprocessedExcelFiles(folder);
    
    if (unprocessedFiles.length === 0) {
      Logger.log("📭 No unprocessed Excel files found. Exiting.");
      return;
    }
    
    Logger.log(`📊 Found ${unprocessedFiles.length} unprocessed file(s)`);
    
    // Step 3: Load reference data once
    Logger.log("📚 Loading reference data...");
    const serialNumberMap = loadSerialNumberMap();
    const schedulingData = loadEquipmentSchedulingData();
    const prepBayData = loadPrepBayAssignments();
    
    Logger.log(`✅ Loaded ${serialNumberMap.size} serial number mappings`);
    Logger.log(`✅ Loaded scheduling data from ${schedulingData.sheets.length} sheets`);
    Logger.log(`✅ Loaded ${prepBayData.length} prep bay assignments`);
    
    // Step 4: Process each file
    const allAlerts = [];
    for (const file of unprocessedFiles) {
      try {
        Logger.log(`\n📄 Processing file: ${file.getName()}`);
        const fileAlerts = processF2File(file, folder, serialNumberMap, schedulingData, prepBayData);
        allAlerts.push(...fileAlerts);
        Logger.log(`✅ Successfully processed: ${file.getName()}`);
      } catch (error) {
        Logger.log(`❌ Error processing ${file.getName()}: ${error.toString()}`);
        Logger.log(`Stack trace: ${error.stack}`);
        // Continue with next file even if one fails
      }
    }
    
    // Step 5: Write alerts if any
    if (allAlerts.length > 0) {
      writeAlerts(allAlerts);
    }
    
    Logger.log("\n✅ F2 Import and Process completed");
    
  } catch (error) {
    Logger.log(`❌ Error in processF2Imports: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Finds unprocessed Excel files in the folder
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder to search
 * @returns {Array<GoogleAppsScript.Drive.File>} Array of unprocessed Excel files
 */
function findUnprocessedExcelFiles(folder) {
  const unprocessedFiles = [];
  const processedFiles = getProcessedFilesList();
  
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Check if file matches the naming pattern
    if (!FILE_NAME_PATTERN.test(fileName)) {
      Logger.log(`⚠️ Skipping file (doesn't match pattern): ${fileName}`);
      continue;
    }
    
    // Check if file has already been processed
    if (processedFiles.has(fileName)) {
      Logger.log(`✓ File already processed: ${fileName}`);
      continue;
    }
    
    unprocessedFiles.push(file);
  }
  
  return unprocessedFiles;
}

/**
 * Processes a single F2 Excel file
 * @param {GoogleAppsScript.Drive.File} file - The Excel file to process
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder containing the file
 * @param {Map} serialNumberMap - Map of barcode to serial number
 * @param {Object} schedulingData - Equipment scheduling data
 * @param {Array} prepBayData - Prep bay assignment data
 * @returns {Array} Array of alert objects
 */
function processF2File(file, folder, serialNumberMap, schedulingData, prepBayData) {
  const alerts = [];
  Logger.log(`📊 File size: ${file.getSize()} bytes`);
  
  // Step 1: Convert Excel to Google Sheets
  Logger.log("🔄 Converting Excel to Google Sheets...");
  const convertedFile = convertExcelToSheets(file);
  
  if (!convertedFile) {
    throw new Error("Failed to convert Excel file to Google Sheets");
  }
  
  Logger.log(`✅ Converted to Google Sheet: ${convertedFile.id}`);
  
  // Step 2: Wait for conversion to complete
  waitForSheetReady(convertedFile.id);
  
  // Step 3: Read and filter the data
  Logger.log("📖 Reading converted sheet data...");
  const serviceData = readF2Data(convertedFile.id);
  
  Logger.log(`📊 Found ${serviceData.length} total service records`);
  
  // Step 4: Filter by equipment category
  const filteredData = serviceData.filter(record => {
    const category = record.EquipmentCategory_lu || '';
    return VALID_EQUIPMENT_CATEGORIES.includes(category);
  });
  
  Logger.log(`📊 Filtered to ${filteredData.length} camera records`);
  
  if (filteredData.length === 0) {
    Logger.log("⚠️ No camera records found in this file");
    // Still mark as processed and move file
    markFileAsProcessed(file.getName());
    moveFileToProcessed(file, folder);
    DriveApp.getFileById(convertedFile.id).setTrashed(true);
    return alerts;
  }
  
  // Step 5: Add Serial Numbers and verification
  Logger.log("🔍 Adding Serial Numbers and verifying data...");
  const enrichedData = [];
  
  for (const record of filteredData) {
    const barcode = record.AssetBarcode ? record.AssetBarcode.toString().trim() : '';
    const orderNumber = record.OrderNumber_lu ? record.OrderNumber_lu.toString().trim() : '';
    
    // Lookup serial number
    const serialNumber = serialNumberMap.get(barcode) || '';
    record.SerialNumber = serialNumber;
    
    // Verify against Equipment Scheduling Chart
    const verification = verifyAgainstScheduling(barcode, orderNumber, schedulingData);
    record.VerificationStatus = verification.status;
    record.VerificationNotes = verification.notes;
    
    // Secondary check against Prep Bay Assignments
    const prepBayCheck = verifyAgainstPrepBay(orderNumber, prepBayData);
    if (prepBayCheck.found && verification.status !== 'Verified') {
      record.VerificationStatus = 'Verified (Prep Bay)';
      record.VerificationNotes = verification.notes + ' | Prep Bay match found';
    }
    
    // Add import metadata
    record.ImportDate = new Date();
    record.ImportTimestamp = new Date().toISOString();
    record.SourceFile = file.getName();
    
    enrichedData.push(record);
    
    // Collect alerts for mismatches
    if (verification.status !== 'Verified' && verification.status !== 'Verified (Prep Bay)') {
      alerts.push({
        barcode: barcode,
        serialNumber: serialNumber,
        orderNumber: orderNumber,
        equipmentName: record.EquipmentName_lu || '',
        issue: verification.status,
        notes: verification.notes,
        sourceFile: file.getName()
      });
    }
  }
  
  // Step 6: Write to F2 Imports sheet
  Logger.log("💾 Writing data to F2 Imports sheet...");
  writeToF2ImportsSheet(enrichedData);
  
  // Step 7: Update RTR Database with Order #s
  Logger.log("🔄 Updating RTR Database with Order #s...");
  updateRTRDatabase(enrichedData);
  
  // Step 8: Mark file as processed and move to Processed folder
  markFileAsProcessed(file.getName());
  moveFileToProcessed(file, folder);
  
  // Step 9: Clean up converted sheet
  Logger.log("🗑️ Cleaning up temporary converted sheet...");
  DriveApp.getFileById(convertedFile.id).setTrashed(true);
  
  Logger.log(`✅ Processed ${enrichedData.length} records, ${alerts.length} alerts generated`);
  return alerts;
}

/**
 * Converts an Excel file to Google Sheets format
 * @param {GoogleAppsScript.Drive.File} file - The Excel file
 * @returns {Object|null} The converted file object with id property, or null if failed
 */
function convertExcelToSheets(file) {
  try {
    // Check if Drive API is available
    if (typeof Drive === 'undefined') {
      throw new Error('Drive API v2 is not enabled. Please enable it in Apps Script: Extensions > Apps Script > Services > + Add Service > Google Drive API v2');
    }
    
    // Initial wait before conversion for large files
    if (file.getSize() > 1000000) { // If file is larger than 1MB
      Logger.log("⏳ Large file detected, waiting 10 seconds before conversion...");
      Utilities.sleep(10000);
    }
    
    // Convert uploaded file to Google Sheets
    const convertedFile = Drive.Files.copy({
      title: `F2_Import_${file.getName().replace('.xlsx', '')}_${new Date().getTime()}`,
      mimeType: MimeType.GOOGLE_SHEETS
    }, file.getId());
    
    return convertedFile;
    
  } catch (error) {
    Logger.log(`❌ Error converting file: ${error.toString()}`);
    throw error;
  }
}

/**
 * Waits for a converted sheet to be ready for reading
 * @param {string} sheetId - The ID of the converted sheet
 */
function waitForSheetReady(sheetId) {
  let ready = false;
  let attempts = 0;
  const maxAttempts = 30;
  const waitTime = 10000; // 10 seconds
  
  // Initial wait after conversion
  Logger.log("⏳ Waiting 20 seconds for initial conversion processing...");
  Utilities.sleep(20000);
  
  while (!ready && attempts++ < maxAttempts) {
    try {
      // Try to open the sheet to verify it's really ready
      const testSheet = SpreadsheetApp.openById(sheetId);
      const testRange = testSheet.getSheets()[0].getRange("A1").getValue();
      ready = true;
      Logger.log("✅ File is ready and accessible.");
    } catch (e) {
      Logger.log(`⚠️ Attempt ${attempts}: File not yet accessible: ${e.toString()}`);
      if (attempts < maxAttempts) {
        Logger.log(`⏳ Waiting ${waitTime/1000} seconds before next attempt...`);
        Utilities.sleep(waitTime);
      }
    }
  }
  
  if (!ready) {
    throw new Error(`❌ Conversion timeout: File not ready after ${(maxAttempts * waitTime)/1000} seconds.`);
  }
}

/**
 * Reads data from the converted F2 sheet
 * @param {string} sheetId - The ID of the converted Google Sheet
 * @returns {Array<Object>} Array of service record objects
 */
function readF2Data(sheetId) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId);
    const dataSheet = sheet.getSheets()[0];
    const dataRange = dataSheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      Logger.log("⚠️ No data rows found (only header row)");
      return [];
    }
    
    // Extract headers (row 1)
    const headers = values[0];
    
    // Map header names to indices
    const headerMap = {};
    headers.forEach((header, index) => {
      if (header) {
        headerMap[header.toString().trim()] = index;
      }
    });
    
    Logger.log(`📋 Headers found: ${Object.keys(headerMap).join(', ')}`);
    
    // Process data rows (skip header row)
    const serviceRecords = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Skip completely empty rows
      if (row.every(cell => !cell || cell.toString().trim() === '')) {
        continue;
      }
      
      // Create record object
      const record = {};
      Object.keys(headerMap).forEach(header => {
        const colIndex = headerMap[header];
        record[header] = row[colIndex] || '';
      });
      
      serviceRecords.push(record);
    }
    
    return serviceRecords;
    
  } catch (error) {
    Logger.log(`❌ Error reading F2 data: ${error.toString()}`);
    throw error;
  }
}

/**
 * Loads Serial Number map from Barcode & Serial Database sheet
 * @returns {Map} Map of barcode (string) to serial number (string)
 */
function loadSerialNumberMap() {
  try {
    const spreadsheet = SpreadsheetApp.openById(F2_DESTINATION_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(BARCODE_SERIAL_SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`⚠️ Sheet "${BARCODE_SERIAL_SHEET_NAME}" not found`);
      return new Map();
    }
    
    const data = sheet.getDataRange().getValues();
    const serialMap = new Map();
    
    // Column H = Serial Number (index 7)
    // Column I = Barcode (index 8)
    for (let i = 1; i < data.length; i++) { // Skip header row
      const barcode = data[i][8] ? data[i][8].toString().trim() : '';
      const serialNumber = data[i][7] ? data[i][7].toString().trim() : '';
      
      if (barcode && serialNumber) {
        serialMap.set(barcode, serialNumber);
      }
    }
    
    Logger.log(`📚 Loaded ${serialMap.size} barcode-to-serial mappings`);
    return serialMap;
    
  } catch (error) {
    Logger.log(`❌ Error loading serial number map: ${error.toString()}`);
    return new Map();
  }
}

/**
 * Loads Equipment Scheduling Chart data
 * @returns {Object} Object containing scheduling data from all relevant sheets
 */
function loadEquipmentSchedulingData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(EQUIPMENT_SCHEDULING_CHART_ID);
    const cameraSheet = spreadsheet.getSheetByName('Camera');
    
    if (!cameraSheet) {
      Logger.log(`⚠️ Camera sheet not found in Equipment Scheduling Chart`);
      return { sheets: [], data: [] };
    }
    
    // Get all data from Camera sheet
    const data = cameraSheet.getDataRange().getValues();
    
    // Extract barcode-order mappings
    // Barcodes are in column E (index 4) with format "BC# [barcode]"
    // Order numbers appear in columns F and beyond (index 5+)
    const barcodeOrderMap = new Map(); // key: "barcode|order", value: true
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const barcodeCell = row[4]; // Column E
      
      // Extract barcode from column E
      let barcode = '';
      if (typeof barcodeCell === 'string') {
        const match = barcodeCell.match(/BC#\s*([A-Z0-9-]+)/);
        if (match) {
          barcode = match[1];
        }
      }
      
      if (!barcode) continue;
      
      // Look for order numbers in columns F and beyond
      for (let colIdx = 5; colIdx < row.length; colIdx++) {
        const cellValue = row[colIdx];
        if (cellValue && typeof cellValue === 'string') {
          // Extract 6-digit order numbers
          const orderMatches = cellValue.match(/\b\d{6}\b/g);
          if (orderMatches) {
            orderMatches.forEach(order => {
              const key = `${barcode}|${order}`;
              barcodeOrderMap.set(key, true);
            });
          }
        }
      }
    }
    
    Logger.log(`📚 Loaded ${barcodeOrderMap.size} barcode-order pairs from Equipment Scheduling Chart`);
    
    return {
      sheets: ['Camera'],
      data: barcodeOrderMap // This is a Map, not an object
    };
    
  } catch (error) {
    Logger.log(`❌ Error loading Equipment Scheduling data: ${error.toString()}`);
    return { sheets: [], data: new Map() };
  }
}

/**
 * Loads Prep Bay Assignments data
 * @returns {Array} Array of prep bay assignment objects
 */
function loadPrepBayAssignments() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PREP_BAY_SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    const prepBayData = [];
    const today = new Date();
    const sevenDaysFromNow = new Date();
    sevenDaysFromNow.setDate(today.getDate() + 7);
    
    // Process visible sheets from today to 7 days out
    const visibleSheets = sheets.filter(sheet => !sheet.isSheetHidden());
    
    for (const sheet of visibleSheets) {
      const sheetName = sheet.getName();
      const dateMatch = sheetName.match(/\w+ (\d+)\/(\d+)/);
      
      if (!dateMatch) continue;
      
      const month = parseInt(dateMatch[1], 10) - 1;
      const day = parseInt(dateMatch[2], 10);
      const sheetDate = new Date(today.getFullYear(), month, day);
      
      if (sheetDate < today) {
        sheetDate.setFullYear(today.getFullYear() + 1);
      }
      
      if (sheetDate < today || sheetDate > sevenDaysFromNow) continue;
      
      // Get data from columns B:J (Job in B, Order in C)
      const data = sheet.getRange('B:J').getValues();
      
      for (const row of data) {
        const jobName = row[0];
        const orderNumber = row[1] ? row[1].toString().trim().replace(/[^0-9]/g, '') : '';
        
        if (jobName && orderNumber) {
          prepBayData.push({
            jobName: jobName.toString().trim(),
            orderNumber: orderNumber,
            sheetName: sheetName,
            date: sheetDate
          });
        }
      }
    }
    
    Logger.log(`📚 Loaded ${prepBayData.length} prep bay assignments`);
    return prepBayData;
    
  } catch (error) {
    Logger.log(`❌ Error loading Prep Bay Assignments: ${error.toString()}`);
    return [];
  }
}

/**
 * Verifies barcode and order against Equipment Scheduling Chart
 * @param {string} barcode - The barcode to verify
 * @param {string} orderNumber - The order number to verify
 * @param {Object} schedulingData - The scheduling data object
 * @returns {Object} Verification result with status and notes
 */
function verifyAgainstScheduling(barcode, orderNumber, schedulingData) {
  if (!barcode || !orderNumber) {
    return {
      status: 'Missing Data',
      notes: `Barcode: ${barcode || 'missing'}, Order: ${orderNumber || 'missing'}`
    };
  }
  
  const key = `${barcode}|${orderNumber}`;
  const found = schedulingData.data.has(key);
  
  if (found) {
    return {
      status: 'Verified',
      notes: 'Found in Equipment Scheduling Chart'
    };
  } else {
    return {
      status: 'Not Found in Scheduling',
      notes: `Barcode ${barcode} not scheduled for Order ${orderNumber} in Equipment Scheduling Chart`
    };
  }
}

/**
 * Verifies order number against Prep Bay Assignments
 * @param {string} orderNumber - The order number to verify
 * @param {Array} prepBayData - Array of prep bay assignments
 * @returns {Object} Verification result
 */
function verifyAgainstPrepBay(orderNumber, prepBayData) {
  if (!orderNumber) {
    return { found: false };
  }
  
  const normalizedOrder = orderNumber.replace(/[^0-9]/g, '');
  const found = prepBayData.some(assignment => 
    assignment.orderNumber.replace(/[^0-9]/g, '') === normalizedOrder
  );
  
  return {
    found: found,
    assignment: found ? prepBayData.find(a => a.orderNumber.replace(/[^0-9]/g, '') === normalizedOrder) : null
  };
}

/**
 * Writes data to F2 Imports sheet, updating existing records by barcode
 * @param {Array<Object>} data - Array of enriched service records
 */
function writeToF2ImportsSheet(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(F2_DESTINATION_SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(F2_IMPORTS_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(F2_IMPORTS_SHEET_NAME);
      Logger.log(`✅ Created new sheet: ${F2_IMPORTS_SHEET_NAME}`);
    }
    
    // Get existing data if any
    const existingData = sheet.getDataRange().getValues();
    const existingHeaders = existingData.length > 0 ? existingData[0] : [];
    
    // Build header row from first record (includes all original headers + new columns)
    const firstRecord = data[0];
    const headers = Object.keys(firstRecord);
    
    // Write headers if sheet is empty
    if (existingData.length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log(`📋 Wrote headers: ${headers.join(', ')}`);
    }
    
    // Create a map of existing records by barcode (for updates)
    const existingByBarcode = new Map();
    if (existingData.length > 1) {
      const barcodeIndex = headers.indexOf('AssetBarcode');
      if (barcodeIndex >= 0) {
        for (let i = 1; i < existingData.length; i++) {
          const barcode = existingData[i][barcodeIndex] ? existingData[i][barcodeIndex].toString().trim() : '';
          if (barcode) {
            existingByBarcode.set(barcode, i + 1); // Store row number (1-based)
          }
        }
      }
    }
    
    // Prepare data rows
    const rowsToWrite = [];
    const rowsToUpdate = [];
    
    for (const record of data) {
      const barcode = record.AssetBarcode ? record.AssetBarcode.toString().trim() : '';
      const rowData = headers.map(header => record[header] || '');
      
      if (barcode && existingByBarcode.has(barcode)) {
        // Update existing row
        const rowNum = existingByBarcode.get(barcode);
        rowsToUpdate.push({ row: rowNum, data: rowData });
      } else {
        // New row to append
        rowsToWrite.push(rowData);
      }
    }
    
    // Update existing rows
    for (const update of rowsToUpdate) {
      sheet.getRange(update.row, 1, 1, headers.length).setValues([update.data]);
    }
    
    // Append new rows
    if (rowsToWrite.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
    }
    
    Logger.log(`💾 Updated ${rowsToUpdate.length} existing records, added ${rowsToWrite.length} new records`);
    
  } catch (error) {
    Logger.log(`❌ Error writing to F2 Imports sheet: ${error.toString()}`);
    throw error;
  }
}

/**
 * Writes alerts to an Alerts section in the F2 Imports sheet
 * @param {Array<Object>} alerts - Array of alert objects
 */
function writeAlerts(alerts) {
  try {
    const spreadsheet = SpreadsheetApp.openById(F2_DESTINATION_SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(F2_IMPORTS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log("⚠️ F2 Imports sheet not found, cannot write alerts");
      return;
    }
    
    // Find or create Alerts section (look for "Alerts" header)
    let alertsStartRow = null;
    const lastRow = sheet.getLastRow();
    
    for (let i = 1; i <= lastRow; i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      if (cellValue && cellValue.toString().trim().toLowerCase() === 'alerts') {
        alertsStartRow = i;
        break;
      }
    }
    
    // If no Alerts section found, create it after the data
    if (!alertsStartRow) {
      alertsStartRow = sheet.getLastRow() + 3; // Add some spacing
      sheet.getRange(alertsStartRow, 1).setValue('Alerts').setFontWeight('bold');
      alertsStartRow += 1;
      
      // Write alert headers
      const alertHeaders = ['Timestamp', 'Barcode', 'Serial Number', 'Order Number', 'Equipment Name', 'Issue', 'Notes', 'Source File'];
      sheet.getRange(alertsStartRow, 1, 1, alertHeaders.length).setValues([alertHeaders]).setFontWeight('bold');
      alertsStartRow += 1;
    } else {
      // Clear existing alerts (keep headers)
      const alertsEndRow = sheet.getLastRow();
      if (alertsEndRow > alertsStartRow) {
        sheet.getRange(alertsStartRow + 1, 1, alertsEndRow - alertsStartRow, 8).clearContent();
      }
      alertsStartRow += 1; // Move past header row
    }
    
    // Write alerts
    if (alerts.length > 0) {
      const alertRows = alerts.map(alert => [
        new Date(),
        alert.barcode,
        alert.serialNumber,
        alert.orderNumber,
        alert.equipmentName,
        alert.issue,
        alert.notes,
        alert.sourceFile
      ]);
      
      sheet.getRange(alertsStartRow, 1, alertRows.length, 8).setValues(alertRows);
      Logger.log(`⚠️ Wrote ${alerts.length} alerts to F2 Imports sheet`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error writing alerts: ${error.toString()}`);
  }
}

/**
 * Updates RTR Database with Order #s from F2 imports
 * Matches by barcode (primary) or serial number (fallback) and updates Order # column
 * @param {Array<Object>} data - Array of enriched service records
 */
function updateRTRDatabase(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(F2_DESTINATION_SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    
    // Find all Status sheets (sheets containing "Status" in name)
    const statusSheets = allSheets.filter(sheet => 
      sheet.getName().toLowerCase().includes('status')
    );
    
    if (statusSheets.length === 0) {
      Logger.log("⚠️ No Status sheets found in RTR Database");
      return;
    }
    
    Logger.log(`🔄 Found ${statusSheets.length} Status sheets to update`);
    
    // Create a map of barcode/serial to order number from F2 data
    const barcodeToOrder = new Map();
    const serialToOrder = new Map();
    
    for (const record of data) {
      const barcode = record.AssetBarcode ? record.AssetBarcode.toString().trim() : '';
      const serial = record.SerialNumber ? record.SerialNumber.toString().trim() : '';
      const orderNumber = record.OrderNumber_lu ? record.OrderNumber_lu.toString().trim() : '';
      
      if (barcode && orderNumber) {
        barcodeToOrder.set(barcode, orderNumber);
      }
      if (serial && orderNumber) {
        serialToOrder.set(serial, orderNumber);
      }
    }
    
    Logger.log(`📚 Created maps: ${barcodeToOrder.size} barcode-order pairs, ${serialToOrder.size} serial-order pairs`);
    
    let totalUpdated = 0;
    
    // Process each Status sheet
    for (const sheet of statusSheets) {
      const sheetName = sheet.getName();
      Logger.log(`\n📋 Processing sheet: ${sheetName}`);
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 3) {
        Logger.log(`  ⚠️ Sheet has insufficient data (minimum row 3 required)`);
        continue;
      }
      
      // Get all data starting from row 3
      // Based on Update Camera's Location: Column C = Serial (index 2), Column D = Barcode (index 3)
      const data = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).getValues();
      
      // Find Order Number column - look for header containing "Order" (case-insensitive)
      // Check row 1 and row 2 for headers
      let orderColumnIndex = null;
      const headerRow1 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const headerRow2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      for (let i = 0; i < headerRow1.length; i++) {
        const header = headerRow1[i] ? headerRow1[i].toString().toLowerCase() : '';
        if (header.includes('order')) {
          orderColumnIndex = i; // 0-based
          Logger.log(`  📍 Found Order column in row 1: Column ${String.fromCharCode(65 + i)} (index ${i})`);
          break;
        }
      }
      
      if (orderColumnIndex === null) {
        for (let i = 0; i < headerRow2.length; i++) {
          const header = headerRow2[i] ? headerRow2[i].toString().toLowerCase() : '';
          if (header.includes('order')) {
            orderColumnIndex = i;
            Logger.log(`  📍 Found Order column in row 2: Column ${String.fromCharCode(65 + i)} (index ${i})`);
            break;
          }
        }
      }
      
      if (orderColumnIndex === null) {
        Logger.log(`  ⚠️ Order Number column not found in sheet headers, skipping updates`);
        continue;
      }
      
      // Process each row and update Order # if match found
      const updates = [];
      let sheetUpdated = 0;
      
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const barcode = row[3] ? row[3].toString().trim() : ''; // Column D (index 3)
        const serial = row[2] ? row[2].toString().trim() : ''; // Column C (index 2)
        const currentOrder = row[orderColumnIndex] ? row[orderColumnIndex].toString().trim() : '';
        
        let newOrder = null;
        let matchType = '';
        
        // Try barcode first, then serial
        if (barcode && barcodeToOrder.has(barcode)) {
          newOrder = barcodeToOrder.get(barcode);
          matchType = 'barcode';
        } else if (serial && serialToOrder.has(serial)) {
          newOrder = serialToOrder.get(serial);
          matchType = 'serial';
        }
        
        // Update if we found a match and it's different from current
        if (newOrder && newOrder !== currentOrder) {
          const rowNum = i + 3; // +3 because data starts at row 3
          updates.push({
            row: rowNum,
            column: orderColumnIndex + 1, // Convert to 1-based
            value: newOrder,
            barcode: barcode,
            serial: serial,
            matchType: matchType
          });
          sheetUpdated++;
        }
      }
      
      // Apply updates in bulk
      if (updates.length > 0) {
        for (const update of updates) {
          sheet.getRange(update.row, update.column).setValue(update.value);
        }
        Logger.log(`  ✅ Updated ${updates.length} records with Order #s`);
        totalUpdated += updates.length;
      } else {
        Logger.log(`  ℹ️ No updates needed for this sheet`);
      }
    }
    
    Logger.log(`\n✅ RTR Database update complete: ${totalUpdated} total records updated across ${statusSheets.length} sheets`);
    
  } catch (error) {
    Logger.log(`❌ Error updating RTR Database: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    // Don't throw - continue processing even if RTR update fails
  }
}

/**
 * Moves a processed file to the "Processed" subfolder
 * @param {GoogleAppsScript.Drive.File} file - The file to move
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - The parent folder
 */
function moveFileToProcessed(file, parentFolder) {
  try {
    // Find or create "Processed" subfolder
    let processedFolder = null;
    const folders = parentFolder.getFoldersByName('Processed');
    
    if (folders.hasNext()) {
      processedFolder = folders.next();
    } else {
      processedFolder = parentFolder.createFolder('Processed');
      Logger.log("📁 Created 'Processed' subfolder");
    }
    
    // Move file to Processed folder
    file.moveTo(processedFolder);
    Logger.log(`📦 Moved file to Processed folder: ${file.getName()}`);
    
  } catch (error) {
    Logger.log(`❌ Error moving file to Processed folder: ${error.toString()}`);
    // Don't throw - file processing should continue even if move fails
  }
}

/**
 * Gets the list of already processed files
 * Uses PropertiesService to track processed files
 * @returns {Set<string>} Set of processed file names
 */
function getProcessedFilesList() {
  const properties = PropertiesService.getScriptProperties();
  const processedFilesJson = properties.getProperty('F2_IMPORT_PROCESSED_FILES') || '[]';
  const processedFiles = JSON.parse(processedFilesJson);
  return new Set(processedFiles);
}

/**
 * Marks a file as processed
 * @param {string} fileName - Name of the file to mark as processed
 */
function markFileAsProcessed(fileName) {
  const properties = PropertiesService.getScriptProperties();
  const processedFilesJson = properties.getProperty('F2_IMPORT_PROCESSED_FILES') || '[]';
  const processedFiles = JSON.parse(processedFilesJson);
  
  if (!processedFiles.includes(fileName)) {
    processedFiles.push(fileName);
    properties.setProperty('F2_IMPORT_PROCESSED_FILES', JSON.stringify(processedFiles));
    Logger.log(`✓ Marked as processed: ${fileName}`);
  }
}

/**
 * Resets the processed files list (for testing/debugging)
 */
function resetProcessedFilesList() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('F2_IMPORT_PROCESSED_FILES');
  Logger.log("🔄 Processed files list has been reset");
}

