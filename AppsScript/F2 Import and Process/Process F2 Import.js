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
 * 7. Writes to "F2 Imports" sheet with deduplication (Barcode + Creation Timestamp)
 * 8. Creates alerts for mismatches
 * 9. Moves processed files to "Processed" subfolder
 * 10. Updates RTR Database with Order #s
 * 
 * File naming pattern: "Service [date] at [time].xlsx"
 * Example: "Service 2025-12-03 at 3.41.39 PM.xlsx"
 */

// Configuration
const F2_IMPORT_FOLDER_ID = '1nUy7lWNr1BVCAxyLsnFASCTszQpjgEnd';
const FILE_NAME_PATTERN = /^Service \d{4}-\d{2}-\d{2} at \d{1,2}\.\d{2}(\.\d{2})? (AM|PM)\.xlsx$/i;

// Spreadsheet IDs
const F2_DESTINATION_SPREADSHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
const EQUIPMENT_SCHEDULING_CHART_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';
const PREP_BAY_SPREADSHEET_ID = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';

// Equipment categories to process
const VALID_EQUIPMENT_CATEGORIES = ['Digital Cameras', '35mm Cameras', '16mm Cameras'];

// Sheet names
const F2_IMPORTS_SHEET_NAME = 'F2 Imports';
const BARCODE_SERIAL_SHEET_NAME = 'Barcode & Serial Database';

// Header mapping: Original F2 Excel headers -> Common database names
// Row 1: Original Excel headers (18 columns + Serial Number, same for Complete and Incomplete exports)
// Row 2: Common database header names (mapped below)
// Schema for Row 2 (common names) - Column layout in F2 Imports sheet:
//   Column A (1): PrepDate -> Prep Date
//   Column B (2): ServicePriority_aet -> Service Priority
//   Column C (3): AssetBarcode -> Barcode
//   Column D (4): SerialNumber -> Serial Number (populated from Barcode & Serial Database lookup)
//   Column E (5): EquipmentName_lu -> Equipment Name
//   Column F (6): EquipmentCategory_lu -> Camera Type
//   Column G (7): OrderNumber_lu -> Order Number
//   Column H (8): JobName_lu -> Job Name
//   Column I (9): Puller_lu -> Puller
//   Column J (10): PrepTech_lu -> Prep Tech
//   Column K (11): ServiceTech -> Service Tech
//   Column L (12): EstimatedCompletionTime_t -> Estimated Completion Time
//   Column M (13): z_log_CreateHost_ts_ae -> Created Timestamp
//   Column N (14): TimestampStart_ts -> Start Timestamp
//   Column O (15): TimestampEnd_ts -> End Timestamp
//   Column P (16): TimestampDuration_cti -> Duration
//   Column Q (17): ServiceStatus_ct -> Service Status
//   Column R (18): PrepKind_lu -> Prep Kind
//   Column S (19): ServiceNotes -> Service Notes
const HEADER_MAPPING = {
  'PrepDate': 'Prep Date',
  'ServicePriority_aet': 'Service Priority',
  'AssetBarcode': 'Barcode',
  'EquipmentName_lu': 'Equipment Name',
  'EquipmentCategory_lu': 'Camera Type',
  'OrderNumber_lu': 'Order Number',
  'JobName_lu': 'Job Name',
  'Puller_lu': 'Puller',
  'PrepTech_lu': 'Prep Tech',
  'ServiceTech': 'Service Tech',
  'EstimatedCompletionTime_t': 'Estimated Completion Time',
  'z_log_CreateHost_ts_ae': 'Created Timestamp',
  'TimestampStart_ts': 'Start Timestamp',
  'TimestampEnd_ts': 'End Timestamp',
  'TimestampDuration_cti': 'Duration',
  'ServiceStatus_ct': 'Service Status',
  'PrepKind_lu': 'Prep Kind',
  'ServiceNotes': 'Service Notes',
  // Additional metadata columns (added during import)
  'SerialNumber': 'Serial Number',
  'VerificationStatus': 'Verification Status',
  'VerificationNotes': 'Verification Notes',
  'ImportDate': 'Import Date',
  'ImportTimestamp': 'Import Timestamp',
  'SourceFile': 'Source File'
};

/**
 * Main function to process F2 imports
 * Simplified version: Imports all raw data directly without filtering or verification
 * Can be triggered manually or via time-driven trigger
 */
function processF2Imports() {
  Logger.log("🚀 Starting F2 Import (Simplified - Raw Data Only)");
  
  try {
    // Step 1: Get the F2 Import folder
    Logger.log("📁 Accessing F2 Import folder...");
    const folder = DriveApp.getFolderById(F2_IMPORT_FOLDER_ID);
    Logger.log(`✅ Folder accessed: ${folder.getName()}`);
    
    // Step 1.5: Clean up already-processed files
    Logger.log("🧹 Cleaning up already-processed files...");
    cleanupProcessedFiles(folder);
    
    // Step 2: Find unprocessed Excel files
    Logger.log("🔍 Searching for unprocessed Excel files...");
    const unprocessedFiles = findUnprocessedExcelFiles(folder);
    
    if (unprocessedFiles.length === 0) {
      Logger.log("📭 No unprocessed Excel files found. Exiting.");
      return;
    }
    
    Logger.log(`📊 Found ${unprocessedFiles.length} unprocessed file(s)`);
    
    // Step 3: Process each file (raw import only)
    for (const file of unprocessedFiles) {
      try {
        Logger.log(`\n📄 Processing file: ${file.getName()}`);
        processF2FileRaw(file, folder);
        Logger.log(`✅ Successfully processed: ${file.getName()}`);
      } catch (error) {
        Logger.log(`❌ Error processing ${file.getName()}: ${error.toString()}`);
        Logger.log(`Stack trace: ${error.stack}`);
        // Continue with next file even if one fails
      }
    }
    
    Logger.log("\n✅ F2 Import completed");
    
  } catch (error) {
    Logger.log(`❌ Error in processF2Imports: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Cleans up already-processed files by moving them to Processed folder
 * or deleting them if they're already in Processed
 * @param {GoogleAppsScript.Drive.Folder} folder - The main F2 Import folder
 */
function cleanupProcessedFiles(folder) {
  try {
    // Get or create Processed subfolder
    let processedFolder = null;
    const folders = folder.getFoldersByName('Processed');
    if (folders.hasNext()) {
      processedFolder = folders.next();
    } else {
      processedFolder = folder.createFolder('Processed');
      Logger.log("📁 Created 'Processed' subfolder");
    }
    
    // Get list of filenames in Processed folder
    const processedNames = new Set();
    const pFiles = processedFolder.getFiles();
    while (pFiles.hasNext()) {
      const processedFile = pFiles.next();
      if (processedFile.getMimeType() === MimeType.MICROSOFT_EXCEL) {
        processedNames.add(processedFile.getName());
      }
    }
    
    // Get all Excel files in the main folder
    const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    let movedCount = 0;
    let deletedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Only process files that match the naming pattern
      if (!FILE_NAME_PATTERN.test(fileName)) {
        continue;
      }
      
      // Check if file already exists in Processed folder
      if (processedNames.has(fileName)) {
        // File already exists in Processed folder, delete the duplicate from main folder
        Logger.log(`🗑️ Duplicate file found in Processed folder, removing from main folder: ${fileName}`);
        file.setTrashed(true);
        deletedCount++;
      }
    }
    
    // Also clean up any leftover converted Google Sheet files in Temp subfolder
    const tempFolders = folder.getFoldersByName('Temp');
    let tempFilesDeleted = 0;
    if (tempFolders.hasNext()) {
      const tempFolder = tempFolders.next();
      const tempFiles = tempFolder.getFiles();
      while (tempFiles.hasNext()) {
        const file = tempFiles.next();
        const fileName = file.getName();
        // Check if it's a converted Google Sheet file (starts with "F2_Import_")
        if (fileName.startsWith('F2_Import_') && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          Logger.log(`🗑️ Removing leftover converted sheet file: ${fileName}`);
          file.setTrashed(true);
          tempFilesDeleted++;
        }
      }
    }
    
    if (movedCount > 0 || deletedCount > 0 || tempFilesDeleted > 0) {
      Logger.log(`✅ Cleanup complete: ${movedCount} file(s) moved, ${deletedCount} duplicate(s) removed, ${tempFilesDeleted} temp file(s) deleted`);
    } else {
      Logger.log(`✅ No cleanup needed`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error during cleanup: ${error.toString()}`);
    // Don't throw - continue with processing even if cleanup fails
  }
}

/**
 * Finds unprocessed Excel files in the folder
 * Checks if files exist in the "Processed" subfolder instead of using ScriptProperties
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder to search
 * @returns {Array<GoogleAppsScript.Drive.File>} Array of unprocessed Excel files
 */
function findUnprocessedExcelFiles(folder) {
  const unprocessedFiles = [];
  
  // Get list of filenames currently inside the "Processed" subfolder
  const processedNames = new Set();
  const processedFolders = folder.getFoldersByName('Processed');
  if (processedFolders.hasNext()) {
    const pFolder = processedFolders.next();
    const pFiles = pFolder.getFiles();
    while (pFiles.hasNext()) {
      const processedFile = pFiles.next();
      // Only track Excel files (same type as source files)
      if (processedFile.getMimeType() === MimeType.MICROSOFT_EXCEL) {
        processedNames.add(processedFile.getName());
      }
    }
  }
  
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Check if file matches the naming pattern
    if (!FILE_NAME_PATTERN.test(fileName)) {
      Logger.log(`⚠️ Skipping file (doesn't match pattern): ${fileName}`);
      continue;
    }
    
    // Check if file exists in processed folder
    if (processedNames.has(fileName)) {
      Logger.log(`✓ File already exists in Processed folder: ${fileName}`);
      continue;
    }
    
    unprocessedFiles.push(file);
  }
  
  return unprocessedFiles;
}

/**
 * Processes a single F2 Excel file - RAW DATA ONLY (simplified for speed)
 * @param {GoogleAppsScript.Drive.File} file - The Excel file to process
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder containing the file
 */
function processF2FileRaw(file, folder) {
  let convertedFile = null;
  
  try {
    Logger.log(`📊 File size: ${file.getSize()} bytes`);
    
    // Step 1: Convert Excel to Google Sheets
    Logger.log("🔄 Converting Excel to Google Sheets...");
    convertedFile = convertExcelToSheets(file, folder);
    
    if (!convertedFile) {
      throw new Error("Failed to convert Excel file to Google Sheets");
    }
    
    Logger.log(`✅ Converted to Google Sheet: ${convertedFile.id}`);
    
    // Step 2: Wait for conversion to complete
    waitForSheetReady(convertedFile.id);
    
    // Step 3: Read ALL raw data (no filtering)
    Logger.log("📖 Reading raw data from converted sheet...");
    const rawData = readF2Data(convertedFile.id);
    
    Logger.log(`📊 Found ${rawData.length} total records (importing all)`);
    
    if (rawData.length === 0) {
      Logger.log("⚠️ No data found in this file");
      // Still move file to Processed folder
      moveFileToProcessed(file, folder);
      return;
    }
  
    // Step 4: Load Serial Number map for lookups
    Logger.log("📚 Loading Serial Number map from Barcode & Serial Database...");
    const serialNumberMap = loadSerialNumberMap();
    
    // Step 5: Add basic import metadata and Serial Number lookup
    const dataWithMetadata = rawData.map(record => {
      // Add minimal metadata
      record.ImportDate = new Date();
      record.ImportTimestamp = new Date().toISOString();
      record.SourceFile = file.getName();
      
      // Look up Serial Number from barcode and add to record for column AD
      const barcode = record.AssetBarcode ? record.AssetBarcode.toString().trim() : '';
      if (barcode && serialNumberMap.has(barcode)) {
        record.SerialNumber = serialNumberMap.get(barcode);
      } else {
        record.SerialNumber = ''; // Empty if not found in database
      }
      
      return record;
    });
  
    // Step 6: Write to F2 Imports sheet (raw data with dual headers, deduplication applied)
    Logger.log("💾 Writing raw data to F2 Imports sheet...");
    writeToF2ImportsSheet(dataWithMetadata);
  
    // Step 7: Move file to Processed folder
    moveFileToProcessed(file, folder);
    
    Logger.log(`✅ Processed ${dataWithMetadata.length} raw records`);
    
  } finally {
    // Always clean up converted sheet, even if there was an error
    if (convertedFile && convertedFile.id) {
      try {
        Logger.log("🗑️ Cleaning up temporary converted sheet...");
        const convertedFileObj = DriveApp.getFileById(convertedFile.id);
        convertedFileObj.setTrashed(true);
        Logger.log("✅ Converted sheet cleaned up");
      } catch (cleanupError) {
        Logger.log(`⚠️ Error cleaning up converted sheet: ${cleanupError.toString()}`);
        // Don't throw - cleanup failure shouldn't break the process
      }
    }
  }
}

/**
 * Converts an Excel file to Google Sheets format
 * Creates the converted file in a "Temp" subfolder to keep root folder clean
 * @param {GoogleAppsScript.Drive.File} file - The Excel file
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - The parent folder (Service Board Imports)
 * @returns {Object|null} The converted file object with id property, or null if failed
 */
function convertExcelToSheets(file, parentFolder) {
  try {
    // Check if Drive API is available
    if (typeof Drive === 'undefined') {
      throw new Error('Drive API v2 is not enabled. Please enable it in Apps Script: Extensions > Apps Script > Services > + Add Service > Google Drive API v2');
    }
    
    // Get or create "Temp" subfolder
    let tempFolder = null;
    const folders = parentFolder.getFoldersByName('Temp');
    if (folders.hasNext()) {
      tempFolder = folders.next();
    } else {
      tempFolder = parentFolder.createFolder('Temp');
      Logger.log("📁 Created 'Temp' subfolder");
    }
    
    // Initial wait before conversion for large files
    if (file.getSize() > 1000000) { // If file is larger than 1MB
      Logger.log("⏳ Large file detected, waiting 10 seconds before conversion...");
      Utilities.sleep(10000);
    }
    
    // Convert uploaded file to Google Sheets (initially created in same folder as source)
    const convertedFile = Drive.Files.copy({
      title: `F2_Import_${file.getName().replace('.xlsx', '')}_${new Date().getTime()}`,
      mimeType: MimeType.GOOGLE_SHEETS
    }, file.getId());
    
    // Move the converted file to Temp subfolder
    const convertedFileObj = DriveApp.getFileById(convertedFile.id);
    convertedFileObj.moveTo(tempFolder);
    Logger.log("📦 Converted file moved to Temp subfolder");
    
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
 * Writes data to F2 Imports sheet with deduplication based on Barcode + Creation Timestamp
 * Row 1: Original Excel headers (with Serial Number inserted at column AD)
 * Row 2: Common database header names (mapped)
 * Row 3+: Imported data (with Serial Number populated in column AD from lookup)
 * Column layout: AA=PrepDate, AB=ServicePriority, AC=Barcode, AD=Serial Number, AE=Equipment Name, etc.
 * Deduplication: Records with same AssetBarcode AND z_log_CreateHost_ts_ae are considered duplicates
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
    
    // Data starts at column AA (column 27, 1-indexed)
    const START_COLUMN = 27; // Column AA
    
    // Build header rows from first record
    const firstRecord = data[0];
    const originalHeaders = Object.keys(firstRecord);
    
    // Define the correct order for standard F2 columns
    // This ensures PrepKind_lu is in the correct position (Column R, after ServiceStatus_ct)
    const STANDARD_F2_COLUMN_ORDER = [
      'PrepDate',
      'ServicePriority_aet',
      'AssetBarcode',
      'SerialNumber', // Inserted after AssetBarcode
      'EquipmentName_lu',
      'EquipmentCategory_lu',
      'OrderNumber_lu',
      'JobName_lu',
      'Puller_lu',
      'PrepTech_lu',
      'ServiceTech',
      'EstimatedCompletionTime_t',
      'z_log_CreateHost_ts_ae',
      'TimestampStart_ts',
      'TimestampEnd_ts',
      'TimestampDuration_cti',
      'ServiceStatus_ct',
      'PrepKind_lu', // Column R - NEW column
      'ServiceNotes'
    ];
    
    // Separate standard columns from metadata columns
    const metadataColumns = ['ImportDate', 'ImportTimestamp', 'SourceFile', 
                            'VerificationStatus', 'VerificationNotes'];
    
    // Build ordered headers: standard columns first (in correct order), then metadata
    const orderedHeaders = [];
    const seenHeaders = new Set();
    
    // First, add standard F2 columns in the correct order
    for (const standardHeader of STANDARD_F2_COLUMN_ORDER) {
      if (originalHeaders.includes(standardHeader) || standardHeader === 'SerialNumber') {
        orderedHeaders.push(standardHeader);
        seenHeaders.add(standardHeader);
      }
    }
    
    // Then, add any remaining columns (metadata or unexpected columns) in their original order
    for (const header of originalHeaders) {
      if (!seenHeaders.has(header)) {
        orderedHeaders.push(header);
      }
    }
    
    // Create mapped headers (Row 2) - use mapping if available, otherwise use original
    const mappedHeaders = orderedHeaders.map(header => {
      return HEADER_MAPPING[header] || header;
    });
    
    // Check if headers need to be written (sheet is empty or headers changed)
    // Read existing headers from column AA
    const existingHeaderRow1 = sheet.getRange(1, START_COLUMN, 1, orderedHeaders.length).getValues()[0];
    const existingHeaderRow2 = sheet.getRange(2, START_COLUMN, 1, mappedHeaders.length).getValues()[0];
    
    // Check if first cell is empty - this indicates an empty sheet
    const isSheetEmpty = !existingHeaderRow1[0] || existingHeaderRow1[0].toString().trim() === '';
    
    const needsHeaderSetup = isSheetEmpty || 
                            !arraysEqual(existingHeaderRow1, orderedHeaders) ||
                            !arraysEqual(existingHeaderRow2, mappedHeaders);
    
    if (needsHeaderSetup) {
      // Write Row 1: Original headers (with SerialNumber inserted) starting at column AA
      sheet.getRange(1, START_COLUMN, 1, orderedHeaders.length).setValues([orderedHeaders]);
      Logger.log(`📋 Wrote original headers (Row 1, starting at column AA): ${orderedHeaders.join(', ')}`);
      
      // Write Row 2: Mapped/common headers starting at column AA
      sheet.getRange(2, START_COLUMN, 1, mappedHeaders.length).setValues([mappedHeaders]);
      Logger.log(`📋 Wrote mapped headers (Row 2, starting at column AA): ${mappedHeaders.join(', ')}`);
    }
    
    // Step 1: Load existing records for deduplication
    // Create a Map of existing records: key = barcode|creationTimestamp, value = {startTimestamp, endTimestamp}
    const existingRecords = new Map();
    const lastRow = sheet.getLastRow();
    
    // Find column indices for key fields
    const barcodeColIndex = orderedHeaders.indexOf('AssetBarcode');
    const creationTimestampColIndex = orderedHeaders.indexOf('z_log_CreateHost_ts_ae');
    const startTimestampColIndex = orderedHeaders.indexOf('TimestampStart_ts');
    const endTimestampColIndex = orderedHeaders.indexOf('TimestampEnd_ts');
    
    if (lastRow >= 3 && barcodeColIndex !== -1 && creationTimestampColIndex !== -1) {
      // Get existing data (rows 3 onwards) starting from column AA
      const existingRows = sheet.getRange(3, START_COLUMN, lastRow - 2, orderedHeaders.length).getValues();
      
      for (let i = 0; i < existingRows.length; i++) {
        const row = existingRows[i];
        const barcode = row[barcodeColIndex] ? row[barcodeColIndex].toString().trim() : '';
        const creationTimestamp = row[creationTimestampColIndex] ? row[creationTimestampColIndex].toString().trim() : '';
        
        if (barcode && creationTimestamp) {
          // Create unique key: barcode|creationTimestamp
          const key = `${barcode}|${creationTimestamp}`;
          
          // Get timestamps from existing record
          const startTimestamp = (startTimestampColIndex !== -1 && row[startTimestampColIndex]) 
            ? row[startTimestampColIndex].toString().trim() : '';
          const endTimestamp = (endTimestampColIndex !== -1 && row[endTimestampColIndex]) 
            ? row[endTimestampColIndex].toString().trim() : '';
          
          // Store the record with its timestamps
          existingRecords.set(key, {
            startTimestamp: startTimestamp,
            endTimestamp: endTimestamp
          });
        }
      }
      Logger.log(`📚 Loaded ${existingRecords.size} existing records for deduplication check`);
    }
    
    // Step 2: Filter out duplicates from new data, with timestamp comparison
    const uniqueRecords = [];
    const duplicateCount = { count: 0 };
    const updatedCount = { count: 0 };
    
    for (const record of data) {
      const barcode = record.AssetBarcode ? record.AssetBarcode.toString().trim() : '';
      const creationTimestamp = record.z_log_CreateHost_ts_ae ? record.z_log_CreateHost_ts_ae.toString().trim() : '';
      
      if (!barcode || !creationTimestamp) {
        // If missing key fields, include the record (can't deduplicate without both)
        uniqueRecords.push(record);
        continue;
      }
      
      // Create unique key: barcode|creationTimestamp
      const key = `${barcode}|${creationTimestamp}`;
      
      if (existingRecords.has(key)) {
        // Found a record with same barcode and creation timestamp
        // Check if timestamps differ or are missing
        const existingRecord = existingRecords.get(key);
        const newStartTimestamp = record.TimestampStart_ts ? record.TimestampStart_ts.toString().trim() : '';
        const newEndTimestamp = record.TimestampEnd_ts ? record.TimestampEnd_ts.toString().trim() : '';
        
        // Check if we should allow this record (update/add needed)
        let shouldAllow = false;
        
        // Check Start Timestamp
        if (!existingRecord.startTimestamp && newStartTimestamp) {
          // Existing record has no start timestamp, new one does - allow it
          shouldAllow = true;
        } else if (existingRecord.startTimestamp && newStartTimestamp && 
                   existingRecord.startTimestamp !== newStartTimestamp) {
          // Both have start timestamps but they differ - allow it
          shouldAllow = true;
        }
        
        // Check End Timestamp
        if (!existingRecord.endTimestamp && newEndTimestamp) {
          // Existing record has no end timestamp, new one does - allow it
          shouldAllow = true;
        } else if (existingRecord.endTimestamp && newEndTimestamp && 
                   existingRecord.endTimestamp !== newEndTimestamp) {
          // Both have end timestamps but they differ - allow it
          shouldAllow = true;
        }
        
        if (shouldAllow) {
          // Timestamps differ or missing - allow this record (it's an update)
          uniqueRecords.push(record);
          updatedCount.count++;
          // Update the existing record in the map with new timestamps
          existingRecords.set(key, {
            startTimestamp: newStartTimestamp || existingRecord.startTimestamp,
            endTimestamp: newEndTimestamp || existingRecord.endTimestamp
          });
        } else {
          // Exact duplicate - skip it
          duplicateCount.count++;
        }
      } else {
        // Not a duplicate - add to unique records and mark as seen
        const newStartTimestamp = record.TimestampStart_ts ? record.TimestampStart_ts.toString().trim() : '';
        const newEndTimestamp = record.TimestampEnd_ts ? record.TimestampEnd_ts.toString().trim() : '';
        uniqueRecords.push(record);
        existingRecords.set(key, {
          startTimestamp: newStartTimestamp,
          endTimestamp: newEndTimestamp
        });
      }
    }
    
    if (duplicateCount.count > 0) {
      Logger.log(`🔄 Skipped ${duplicateCount.count} exact duplicate record(s) based on Barcode + Creation Timestamp`);
    }
    if (updatedCount.count > 0) {
      Logger.log(`🔄 Allowed ${updatedCount.count} record(s) with updated timestamps (Start or End Timestamp differs or was missing)`);
    }
    
    // Step 3: Write only unique records
    const rowsToWrite = [];
    
    for (const record of uniqueRecords) {
      const rowData = [];
      for (const header of orderedHeaders) {
        // Write Serial Number value (looked up from Barcode & Serial Database)
        rowData.push(record[header] || '');
      }
      rowsToWrite.push(rowData);
    }
    
    // Append all new unique rows (starting after row 2, or after last data row) at column AA
    if (rowsToWrite.length > 0) {
      const currentLastRow = sheet.getLastRow();
      const startRow = currentLastRow < 2 ? 3 : currentLastRow + 1; // Start at row 3 if sheet only has headers
      sheet.getRange(startRow, START_COLUMN, rowsToWrite.length, orderedHeaders.length).setValues(rowsToWrite);
      Logger.log(`💾 Added ${rowsToWrite.length} new unique record(s) (Column AD populated with Serial Numbers, starting at column AA)`);
    } else {
      Logger.log(`ℹ️ No new unique records to add (all were duplicates)`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error writing to F2 Imports sheet: ${error.toString()}`);
    throw error;
  }
}

/**
 * Helper function to compare arrays
 */
function arraysEqual(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * Writes alerts to an Alerts section in the F2 Imports sheet
 * Alerts section starts in column AA (column 27)
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
    
    const ALERTS_START_COL = 27; // Column AA (1-based: A=1, B=2, ..., AA=27)
    
    // Find or create Alerts section (look for "Alerts" header in column AA)
    let alertsStartRow = null;
    const lastRow = sheet.getLastRow();
    
    for (let i = 1; i <= lastRow; i++) {
      const cellValue = sheet.getRange(i, ALERTS_START_COL).getValue();
      if (cellValue && cellValue.toString().trim().toLowerCase() === 'alerts') {
        alertsStartRow = i;
        break;
      }
    }
    
    // If no Alerts section found, create it starting at row 1 in column AA
    if (!alertsStartRow) {
      alertsStartRow = 1;
      sheet.getRange(alertsStartRow, ALERTS_START_COL).setValue('Alerts').setFontWeight('bold');
      alertsStartRow += 1;
      
      // Write alert headers
      const alertHeaders = ['Timestamp', 'Barcode', 'Serial Number', 'Order Number', 'Equipment Name', 'Issue', 'Notes', 'Source File'];
      sheet.getRange(alertsStartRow, ALERTS_START_COL, 1, alertHeaders.length).setValues([alertHeaders]).setFontWeight('bold');
      alertsStartRow += 1;
    } else {
      // Clear existing alerts (keep headers)
      // alertsStartRow points to the "Alerts" label row, headers are in the next row
      const headerRow = alertsStartRow + 1; // Headers are in the row after "Alerts" label
      const alertsEndRow = sheet.getLastRow();
      if (alertsEndRow > headerRow) {
        // Find how many alert columns exist (check row with headers)
        let alertColCount = 8; // Default to 8 columns
        // Try to detect actual column count by checking for empty cells
        for (let col = ALERTS_START_COL; col < ALERTS_START_COL + 20; col++) {
          const headerCell = sheet.getRange(headerRow, col).getValue();
          if (!headerCell || headerCell.toString().trim() === '') {
            alertColCount = col - ALERTS_START_COL;
            break;
          }
        }
        // Clear data rows (skip the header row)
        sheet.getRange(headerRow + 1, ALERTS_START_COL, alertsEndRow - headerRow, alertColCount).clearContent();
      }
      alertsStartRow = headerRow + 1; // Set to first data row after headers
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
      
      sheet.getRange(alertsStartRow, ALERTS_START_COL, alertRows.length, 8).setValues(alertRows);
      Logger.log(`⚠️ Wrote ${alerts.length} alerts to F2 Imports sheet starting at column AA (row ${alertsStartRow})`);
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
      
      // Process each row to build updates - use batch column write for performance
      // Extract current Order column data into a 1D array
      const orderColumnValues = data.map(row => row[orderColumnIndex] || '');
      let updatesMade = false;
      let sheetUpdated = 0;
      
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const barcode = row[3] ? row[3].toString().trim() : ''; // Column D (index 3)
        const serial = row[2] ? row[2].toString().trim() : ''; // Column C (index 2)
        const currentOrder = orderColumnValues[i] ? orderColumnValues[i].toString().trim() : '';
        
        let newOrder = null;
        
        // Try barcode first, then serial
        if (barcode && barcodeToOrder.has(barcode)) {
          newOrder = barcodeToOrder.get(barcode);
        } else if (serial && serialToOrder.has(serial)) {
          newOrder = serialToOrder.get(serial);
        }
        
        // Update the array in memory if we found a match and it's different
        if (newOrder && newOrder !== currentOrder) {
          orderColumnValues[i] = newOrder; // Update the array in memory
          updatesMade = true;
          sheetUpdated++;
        }
      }
      
      // Write the ENTIRE column back in one go if changes were made
      if (updatesMade) {
        // Map the 1D array back to 2D array for setValues (each value becomes [value])
        const columnData = orderColumnValues.map(val => [val]);
        // Write from row 3 down (data starts at row 3)
        sheet.getRange(3, orderColumnIndex + 1, columnData.length, 1).setValues(columnData);
        Logger.log(`  ✅ Batch updated ${sheetUpdated} record(s) with Order #s`);
        totalUpdated += sheetUpdated;
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
 * Web app entry point for triggering F2 import processing
 * Can be called via HTTP GET request from external scripts (e.g., PowerShell)
 * 
 * Security: Requires a token parameter in the URL to prevent unauthorized access
 * Example URL: https://script.google.com/.../exec?token=YOUR_SECRET_TOKEN
 * 
 * To deploy:
 * 1. In Apps Script editor, click "Deploy" > "New deployment"
 * 2. Select type: "Web app"
 * 3. Execute as: "Me"
 * 4. Who has access: "Anyone" (required for unauthenticated requests from PowerShell)
 * 5. Click "Deploy" and copy the Web app URL
 * 6. Add ?token=YOUR_SECRET_TOKEN to the URL
 * 7. Update F2-Desktop-Monitor.ps1 with the full URL including token
 * 
 * To set your token:
 * 1. Choose a random, hard-to-guess string (e.g., "F2Import2025KeslowKey")
 * 2. Add it to the URL: https://script.google.com/.../exec?token=F2Import2024SecretKey123
 * 3. Update the WEB_APP_TOKEN constant below to match
 * 
 * @param {GoogleAppsScript.Events.DoGetEvent} e - The event object
 * @returns {GoogleAppsScript.Content.TextOutput} Response text
 */
function doGet(e) {
  const logPrefix = "[doGet]";
  Logger.log(`${logPrefix} Request received at ${new Date().toISOString()}`);
  
  try {
    // Security: Check for token parameter
    const WEB_APP_TOKEN = "F2Import2025KeslowKey"; // CHANGE THIS to your own secret token
    const providedToken = e && e.parameter ? e.parameter.token : undefined;
    const hasToken = !!providedToken;
    const tokenValid = hasToken && providedToken === WEB_APP_TOKEN;
    
    Logger.log(`${logPrefix} Token present: ${hasToken}, valid: ${tokenValid}`);
    
    if (!tokenValid) {
      Logger.log(`${logPrefix} ❌ Web app access denied - invalid or missing token`);
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Unauthorized - invalid token",
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    Logger.log(`${logPrefix} 🌐 Web app triggered - scheduling F2 import processing`);
    
    // Return immediately to avoid web app timeout
    // Create a time-driven trigger to run processF2Imports() in the next minute
    // This prevents the web app from timing out on long-running operations
    try {
      // Delete any existing triggers for processF2Imports to avoid duplicates
      const existingTriggers = ScriptApp.getProjectTriggers();
      let deletedCount = 0;
      existingTriggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === 'processF2Imports' && 
            trigger.getEventType() === ScriptApp.EventType.CLOCK) {
          ScriptApp.deleteTrigger(trigger);
          deletedCount++;
        }
      });
      if (deletedCount > 0) {
        Logger.log(`${logPrefix} Removed ${deletedCount} existing processF2Imports trigger(s)`);
      }
      
      // Create a one-time trigger to run in 10 seconds
      ScriptApp.newTrigger('processF2Imports')
        .timeBased()
        .after(10000) // Run in 10 seconds
        .create();
      
      Logger.log(`${logPrefix} ✅ Trigger created - processF2Imports will run in 10 seconds`);
    } catch (triggerError) {
      Logger.log(`${logPrefix} ⚠️ Could not create trigger: ${triggerError.toString()}`);
      Logger.log(`${logPrefix} Running processF2Imports directly (fallback)`);
      processF2Imports();
    }
    
    Logger.log(`${logPrefix} Returning success response`);
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "F2 import processing scheduled successfully",
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log(`${logPrefix} ❌ Error in web app: ${error.toString()}`);
    if (error.stack) {
      Logger.log(`${logPrefix} Stack: ${error.stack}`);
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sets up time-driven trigger to automatically process F2 imports
 * This function should be run once manually to set up the trigger
 * 
 * The trigger will run processF2Imports() every 10 minutes to check for new files
 * This serves as a backup/fallback if the webhook from PowerShell doesn't work
 * 
 * To set up:
 * 1. Run this function once manually in Apps Script editor
 * 2. Authorize permissions if prompted
 * 3. Check Triggers (clock icon) to verify it was created
 * 
 * To remove the trigger:
 * 1. Go to Triggers (clock icon) in Apps Script editor
 * 2. Find the trigger for processF2Imports
 * 3. Click the three dots menu and select "Delete"
 */
function setupF2ImportTrigger() {
  try {
    // Delete existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'processF2Imports') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        Logger.log('Deleted existing trigger for processF2Imports');
      }
    });
    
    // Create a time-driven trigger that runs every 10 minutes
    // This will automatically check for new files in the Google Drive folder
    ScriptApp.newTrigger('processF2Imports')
      .timeBased()
      .everyMinutes(10)  // Check every 10 minutes
      .create();
    
    Logger.log('✅ Time-driven trigger created successfully');
    Logger.log('   Function: processF2Imports');
    Logger.log('   Frequency: Every 10 minutes');
    Logger.log(`   Deleted ${deletedCount} existing trigger(s) before creating new one`);
    Logger.log('');
    Logger.log('📋 Next steps:');
    Logger.log('   1. Check Triggers (clock icon) in Apps Script editor to verify');
    Logger.log('   2. The trigger will automatically run processF2Imports() every 10 minutes');
    Logger.log('   3. This serves as a backup if the PowerShell webhook fails');
    
  } catch (error) {
    Logger.log(`❌ Error setting up trigger: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}
