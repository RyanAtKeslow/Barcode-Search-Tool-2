/**
 * Finds the most recent Assets_GoogleExport file in Google Drive
 * @returns {GoogleAppsScript.Drive.File} The most recent Assets file or null if not found
 */
function findMostRecentAssetsFile() {
  try {
    // Search for files with the naming pattern "Assets_GoogleExport_YYYYMMDD_XXXXX"
    const files = DriveApp.searchFiles('title contains "Assets_GoogleExport_"');
    let mostRecentFile = null;
    let mostRecentDate = null;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Extract date from filename (format: Assets_GoogleExport_YYYYMMDD_XXXXX)
      const dateMatch = fileName.match(/Assets_GoogleExport_(\d{8})_/);
      if (dateMatch) {
        const fileDate = new Date(
          dateMatch[1].substring(0, 4), // year
          parseInt(dateMatch[1].substring(4, 6)) - 1, // month (0-indexed)
          dateMatch[1].substring(6, 8) // day
        );
        
        if (!mostRecentDate || fileDate > mostRecentDate) {
          mostRecentDate = fileDate;
          mostRecentFile = file;
        }
      }
    }
    
    if (mostRecentFile) {
      Logger.log(`📁 Found most recent Assets file: ${mostRecentFile.getName()} (${mostRecentDate.toDateString()})`);
    } else {
      Logger.log("❌ No Assets_GoogleExport files found");
    }
    
    return mostRecentFile;
  } catch (error) {
    Logger.log(`❌ Error finding Assets file: ${error.toString()}`);
    return null;
  }
}

/**
 * Reads data from the Assets file
 * @param {GoogleAppsScript.Drive.File} file The Assets file
 * @returns {Array} Array of data rows
 */
function readAssetsFileData(file) {
  try {
    Logger.log(`📁 Reading file: ${file.getName()}`);
    Logger.log(`📊 File size: ${file.getSize()} bytes`);
    
    let sheetFileId;
    
    // Check if file is already a Google Sheet
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      Logger.log("📊 File is already a Google Sheet. Using directly.");
      sheetFileId = file.getId();
    } else {
      Logger.log("📊 File is not a Google Sheet. Converting...");
      
      // Initial wait before conversion for large files
      if (file.getSize() > 1000000) { // If file is larger than 1MB
        Logger.log("⏳ Large file detected, waiting 10 seconds before conversion...");
        Utilities.sleep(10000);
      }
      
      // Convert file to Google Sheets
      const convertedFile = Drive.Files.copy({
        title: file.getName(),
        mimeType: MimeType.GOOGLE_SHEETS
      }, file.getId());
      
      Logger.log(`📊 Converted to Google Sheet: ${convertedFile.id}`);
      sheetFileId = convertedFile.id;
      
      // Wait until the converted file is fully ready
      let ready = false;
      let attempts = 0;
      const maxAttempts = 30;
      const waitTime = 10000;
      
      // Initial wait after conversion
      Logger.log("⏳ Waiting 20 seconds for initial conversion processing...");
      Utilities.sleep(20000);
      
      while (!ready && attempts++ < maxAttempts) {
        try {
          const fileSize = DriveApp.getFileById(sheetFileId).getSize();
          Logger.log(`⏳ Attempt ${attempts}: File size is ${fileSize} bytes`);
          
          // Try to open the sheet to verify it's really ready
          try {
            const testSheet = SpreadsheetApp.openById(sheetFileId);
            const testRange = testSheet.getSheets()[0].getRange("A1").getValue();
            ready = true;
            Logger.log("✅ File is ready and accessible.");
          } catch (e) {
            Logger.log(`⚠️ File not yet accessible: ${e.toString()}`);
            ready = false;
          }
        } catch (e) {
          Logger.log(`⚠️ Attempt ${attempts} failed to get file size: ${e.toString()}`);
        }
        
        if (!ready) {
          Logger.log(`⏳ Waiting ${waitTime/1000} seconds before next attempt...`);
          Utilities.sleep(waitTime);
        }
      }
      
      if (!ready) {
        throw new Error(`❌ Conversion timeout: File not ready after ${(maxAttempts * waitTime)/1000} seconds.`);
      }
    }
    
    // Open the converted sheet and get data
    Logger.log("🔍 Opening converted sheet...");
    const convertedSheet = SpreadsheetApp.openById(sheetFileId);
    const sourceSheet = convertedSheet.getSheets()[0];
    
    Logger.log(`📊 Source sheet has ${sourceSheet.getLastRow()} rows`);
    
    // Get all data at once using the same method as F2DataDumpDirectPrint
    Logger.log("📊 Reading source data...");
    const data = sourceSheet.getDataRange().getValues();
    
    Logger.log(`📊 Successfully read ${data.length} rows from Assets file`);
    
    // Log first few rows for debugging
    for (let i = 0; i < Math.min(3, data.length); i++) {
      Logger.log(`📋 Row ${i}: ${JSON.stringify(data[i])}`);
    }
    
    return data;
    
  } catch (error) {
    Logger.log(`❌ Error reading Assets file: ${error.toString()}`);
    Logger.log(`❌ Error stack: ${error.stack}`);
    return [];
  }
}

// Standalone function to export data from Assets file to secondary database
function runSecondaryDatabaseExport() {
  Logger.log("🔄 Starting standalone secondary database export...");
  
  try {
    // Find the most recent Assets_GoogleExport file
    Logger.log("🔍 Searching for Assets_GoogleExport file...");
    const assetsFile = findMostRecentAssetsFile();
    
    if (!assetsFile) {
      throw new Error("No Assets_GoogleExport file found in Google Drive");
    }
    
    Logger.log(`📁 Found Assets file: ${assetsFile.getName()}`);
    
    // Read data from the Assets file
    Logger.log("📊 Reading data from Assets file...");
    const allData = readAssetsFileData(assetsFile);
    
    Logger.log(`📊 Assets file data length: ${allData ? allData.length : 'null'}`);
    
    if (!allData || allData.length <= 1) {
      Logger.log("⚠️ No data found in Assets file");
      return {
        success: false,
        message: "No data found in Assets file",
        rowsExported: 0
      };
    }
    
    Logger.log(`📊 Successfully read ${allData.length} rows from Assets file`);
    
    // Target spreadsheet ID from the URL
    const targetSpreadsheetId = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
    const targetSheetName = 'Barcode Database';
    
    // Open the target spreadsheet
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    if (!targetSpreadsheet) {
      throw new Error("Could not access target spreadsheet");
    }
    
    // Get or create the target sheet
    let targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet) {
      targetSheet = targetSpreadsheet.insertSheet(targetSheetName);
      Logger.log(`📄 Created new sheet: ${targetSheetName}`);
    }
    
    // Define the new headers for the secondary export
    const secondaryHeaders = [
      "Category", 
      "Location", 
      "Status", 
      "Equip Name", 
      "Owner", 
      "Equipment ID", 
      "Asset ID", 
      "Serial Number", 
      "Barcode"
    ];
    
    // Format the data for secondary export
    Logger.log("📊 Preparing secondary export data...");
    let formattedData;
    try {
      formattedData = formatDataForSecondaryExport(allData);
      Logger.log(`📊 Successfully formatted ${formattedData.length} rows for export`);
    } catch (formatError) {
      Logger.log(`❌ Error formatting data: ${formatError.toString()}`);
      Logger.log(`❌ Format error stack: ${formatError.stack}`);
      throw formatError;
    }
    
    if (!formattedData || formattedData.length === 0) {
      Logger.log("⚠️ No data to export to secondary database.");
      return {
        success: false,
        message: "No data to export",
        rowsExported: 0,
        targetSheet: targetSheetName,
        targetSpreadsheet: targetSpreadsheet.getName()
      };
    }
    
    // Clear the target sheet
    targetSheet.clearContents();
    targetSheet.clearFormats();
    
    // Add completion timestamp
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM');
    targetSheet.getRange(1, 1).setValue(`Secondary Database Export Completed on ${today}`);
    
    // Add headers starting from row 2
    targetSheet.getRange(2, 1, 1, secondaryHeaders.length).setValues([secondaryHeaders]);
    
    // Write the formatted data starting from row 3
    // Write the formatted data in chunks to prevent timeout
    const chunkSize = 25000; // Process 25000 rows at a time
    const totalRows = formattedData.length;
    const numCols = formattedData[0].length;
    
    Logger.log(`📊 Writing ${totalRows} rows in chunks of ${chunkSize}...`);
    
    for (let i = 0; i < totalRows; i += chunkSize) {
      const endRow = Math.min(i + chunkSize, totalRows);
      const chunk = formattedData.slice(i, endRow);
      const startRow = i + 3; // Data starts from row 3
      
      targetSheet.getRange(startRow, 1, chunk.length, numCols).setValues(chunk);
      
      const progress = Math.round((endRow / totalRows) * 100);
      Logger.log(`📝 Written rows ${i + 1}-${endRow} (${progress}% complete)`);
      
      // Small delay to prevent overwhelming the API
      if (i + chunkSize < totalRows) {
        Utilities.sleep(200); // Slightly longer delay for larger chunks
      }
    }
    
    // Set frozen rows
    targetSheet.setFrozenRows(2);
    
    Logger.log(`✅ Secondary database export completed successfully. Exported ${formattedData.length} rows.`);
    
    // Send completion email
    try {
      MailApp.sendEmail({
        to: "Owen@keslowcamera.com, ryan@keslowcamera.com",
        subject: "✅ Secondary Database Export Completed",
        body: `Secondary database export completed successfully.\n\n` +
          `Source: Barcode Dictionary sheet\n` +
          `Target: ${targetSpreadsheet.getName()} - ${targetSheetName}\n` +
          `Rows exported: ${formattedData.length}\n` +
          `Completed on: ${today}`
      });
      Logger.log('✅ Completion email sent.');
    } catch (emailError) {
      Logger.log(`❌ Error sending completion email: ${emailError.toString()}`);
    }
    
    return {
      success: true,
      message: "Secondary database export completed",
      rowsExported: formattedData.length,
      targetSheet: targetSheetName,
      targetSpreadsheet: targetSpreadsheet.getName()
    };
    
  } catch (error) {
    Logger.log(`❌ Error in secondary database export: ${error.toString()}`);
    
    // Send error email
    try {
      MailApp.sendEmail({
        to: "Owen@keslowcamera.com, ryan@keslowcamera.com",
        subject: "❌ Secondary Database Export Failed",
        body: `Secondary database export failed with error:\n\n${error.toString()}\n\nPlease check the logs for more details.`
      });
    } catch (emailError) {
      Logger.log(`❌ Error sending error email: ${emailError.toString()}`);
    }
    
    return {
      success: false,
      message: error.toString(),
      rowsExported: 0,
      targetSheet: 'Barcode Database',
      targetSpreadsheet: 'Secondary Database'
    };
  }
}

// Legacy function for backward compatibility (if called from Data Dump)
function exportToSecondaryDatabase(processedData, summaryStats) {
  Logger.log("🔄 Starting secondary database export (legacy mode)...");
  
  try {
    // Target spreadsheet ID from the URL
    const targetSpreadsheetId = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
    const targetSheetName = 'Barcode Database';
    
    // Open the target spreadsheet
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    if (!targetSpreadsheet) {
      throw new Error("Could not access target spreadsheet");
    }
    
    // Get or create the target sheet
    let targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet) {
      targetSheet = targetSpreadsheet.insertSheet(targetSheetName);
      Logger.log(`📄 Created new sheet: ${targetSheetName}`);
    }
    
    // Define the new headers for the secondary export
    const secondaryHeaders = [
      "Category", 
      "Location", 
      "Status", 
      "Equip Name", 
      "Owner", 
      "Equipment ID", 
      "Asset ID", 
      "Serial Number", 
      "Barcode"
    ];
    
    // Format the data for secondary export
    Logger.log("📊 Preparing secondary export data...");
    const formattedData = formatDataForSecondaryExport(processedData);
    
    if (!formattedData || formattedData.length === 0) {
      Logger.log("⚠️ No data to export to secondary database.");
      return {
        success: false,
        message: "No data to export",
        rowsExported: 0,
        targetSheet: targetSheetName,
        targetSpreadsheet: targetSpreadsheet.getName()
      };
    }
    
    // Clear the target sheet
    targetSheet.clearContents();
    targetSheet.clearFormats();
    
    // Add completion timestamp
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM');
    targetSheet.getRange(1, 1).setValue(`Secondary Database Export Completed on ${today}`);
    
    // Add headers starting from row 2
    targetSheet.getRange(2, 1, 1, secondaryHeaders.length).setValues([secondaryHeaders]);
    
    // Write the formatted data starting from row 3
    // Write the formatted data in chunks to prevent timeout
    const chunkSize = 25000; // Process 25000 rows at a time
    const totalRows = formattedData.length;
    const numCols = formattedData[0].length;
    
    Logger.log(`📊 Writing ${totalRows} rows in chunks of ${chunkSize}...`);
    
    for (let i = 0; i < totalRows; i += chunkSize) {
      const endRow = Math.min(i + chunkSize, totalRows);
      const chunk = formattedData.slice(i, endRow);
      const startRow = i + 3; // Data starts from row 3
      
      targetSheet.getRange(startRow, 1, chunk.length, numCols).setValues(chunk);
      
      const progress = Math.round((endRow / totalRows) * 100);
      Logger.log(`📝 Written rows ${i + 1}-${endRow} (${progress}% complete)`);
      
      // Small delay to prevent overwhelming the API
      if (i + chunkSize < totalRows) {
        Utilities.sleep(200); // Slightly longer delay for larger chunks
      }
    }
    
    // Set frozen rows
    targetSheet.setFrozenRows(2);
    
    Logger.log(`✅ Secondary database export completed successfully. Exported ${formattedData.length} rows.`);
    
    return {
      success: true,
      message: "Secondary database export completed",
      rowsExported: formattedData.length,
      targetSheet: targetSheetName,
      targetSpreadsheet: targetSpreadsheet.getName()
    };
    
  } catch (error) {
    Logger.log(`❌ Error in secondary database export: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
      rowsExported: 0,
      targetSheet: 'Barcode Database',
      targetSpreadsheet: 'Secondary Database'
    };
  }
}

// Helper function to format data for secondary export
function formatDataForSecondaryExport(rawData) {
  if (!rawData || !rawData.length) return [];
  
  Logger.log("🔄 Formatting data for secondary export...");
  
  // Debug: Log the header row to understand the Assets file schema
  if (rawData.length > 0) {
    Logger.log(`📋 Assets file headers: ${JSON.stringify(rawData[0])}`);
  }
  
  // Map the Assets file schema to our target schema
  // Source schema: __kp_Asset_id, _kf_Equip_id, asset_EQUIP::EquipName_aer, asset_equip_EQUIPCAT::EquipcatName_aer, AssetBarcode, AssetSerialNumber, AssetStatus, AssetOwner_lu, AssetLocation_lu
  // Target schema: Category, Location, Status, Equip Name, Owner, Equipment ID, Asset ID, Serial Number, Barcode
  
  const COLUMNS = {
    ASSET_ID: 0,        // __kp_Asset_id -> Asset ID
    EQUIP_ID: 1,        // _kf_Equip_id -> Equipment ID  
    EQUIP_NAME: 2,      // asset_EQUIP::EquipName_aer -> Equip Name
    CATEGORY: 3,        // asset_equip_EQUIPCAT::EquipcatName_aer -> Category
    BARCODE: 4,         // AssetBarcode -> Barcode
    SERIAL: 5,          // AssetSerialNumber -> Serial Number
    STATUS: 6,          // AssetStatus -> Status
    OWNER: 7,           // AssetOwner_lu -> Owner
    LOCATION: 8         // AssetLocation_lu -> Location
  };
  
  Logger.log(`📋 Using fixed column mapping: ${JSON.stringify(COLUMNS)}`);
  
  const formattedData = [];
  
  // Process each row (skip header row)
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // Debug: Log the first few rows to understand the data structure
    if (i <= 3) {
      Logger.log(`Row ${i} data: ${JSON.stringify(row)}`);
    }
    
    // Get the barcode string and explode it into individual barcodes
    const barcodeString = row[COLUMNS.BARCODE] || '';
    const barcodes = barcodeString.toString().split('|').map(b => b.trim()).filter(b => b.length > 0);
    
    // Debug: Log barcode processing for first few rows
    if (i <= 3) {
      Logger.log(`Row ${i} barcode string: "${barcodeString}"`);
      Logger.log(`Row ${i} exploded barcodes: ${JSON.stringify(barcodes)}`);
    }
    
    // If no barcodes, create one record with empty barcode
    if (barcodes.length === 0) {
      const formattedRow = [
        row[COLUMNS.CATEGORY] || '',           // Category (asset_equip_EQUIPCAT::EquipcatName_aer)
        row[COLUMNS.LOCATION] || '',           // Location (AssetLocation_lu)
        row[COLUMNS.STATUS] || '',             // Status (AssetStatus)
        row[COLUMNS.EQUIP_NAME] || '',         // Equip Name (asset_EQUIP::EquipName_aer)
        row[COLUMNS.OWNER] || '',              // Owner (AssetOwner_lu)
        row[COLUMNS.EQUIP_ID] || '',           // Equipment ID (_kf_Equip_id)
        row[COLUMNS.ASSET_ID] || '',           // Asset ID (__kp_Asset_id)
        row[COLUMNS.SERIAL] || '',             // Serial Number (AssetSerialNumber)
        ''                                     // Barcode (empty)
      ];
      formattedData.push(formattedRow);
    } else {
      // Create one record for each barcode
      for (const barcode of barcodes) {
        const formattedRow = [
          row[COLUMNS.CATEGORY] || '',           // Category (asset_equip_EQUIPCAT::EquipcatName_aer)
          row[COLUMNS.LOCATION] || '',           // Location (AssetLocation_lu)
          row[COLUMNS.STATUS] || '',             // Status (AssetStatus)
          row[COLUMNS.EQUIP_NAME] || '',         // Equip Name (asset_EQUIP::EquipName_aer)
          row[COLUMNS.OWNER] || '',              // Owner (AssetOwner_lu)
          row[COLUMNS.EQUIP_ID] || '',           // Equipment ID (_kf_Equip_id)
          row[COLUMNS.ASSET_ID] || '',           // Asset ID (__kp_Asset_id)
          row[COLUMNS.SERIAL] || '',             // Serial Number (AssetSerialNumber)
          barcode                                // Barcode (individual)
        ];
        formattedData.push(formattedRow);
      }
    }
  }
  
  Logger.log(`✅ Formatted ${formattedData.length} rows for secondary export (exploded barcodes)`);
  return formattedData;
} 