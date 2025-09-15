// Standalone function to export data from Barcode Dictionary to secondary database
function runSecondaryDatabaseExport() {
  Logger.log("🔄 Starting standalone secondary database export...");
  
  try {
    // Get the active spreadsheet (should be the main barcode spreadsheet)
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!sourceSpreadsheet) {
      throw new Error("Could not access active spreadsheet");
    }
    
    // Get the Barcode Dictionary sheet
    const barcodeSheet = sourceSpreadsheet.getSheetByName('Barcode Dictionary');
    if (!barcodeSheet) {
      throw new Error("Barcode Dictionary sheet not found");
    }
    
    // Check if there's data in the sheet (skip header row and completion message)
    const lastRow = barcodeSheet.getLastRow();
    if (lastRow < 3) { // Less than 3 rows means no data (completion message + header + at least 1 data row)
      Logger.log("⚠️ No data found in Barcode Dictionary sheet.");
      return {
        success: false,
        message: "No data found in Barcode Dictionary sheet",
        rowsExported: 0
      };
    }
    
    // Read all data from the Barcode Dictionary sheet (starting from row 2 to skip completion message)
    Logger.log(`📊 Reading data from Barcode Dictionary sheet (${lastRow} rows)...`);
    const allData = barcodeSheet.getRange(2, 1, lastRow - 1, barcodeSheet.getLastColumn()).getValues();
    
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
    const formattedData = formatDataForSecondaryExport(allData);
    
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
    targetSheet.getRange(3, 1, formattedData.length, formattedData[0].length).setValues(formattedData);
    
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
    targetSheet.getRange(3, 1, formattedData.length, formattedData[0].length).setValues(formattedData);
    
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
  
  // Column mapping for the new structure (including previously ignored columns)
  const COLUMNS = {
    ASSET_ID: 0,        // Column A - Asset ID
    UUID: 1,            // Column B
    EQUIPMENT: 2,       // Column C - Equipment Name
    CATEGORY: 3,        // Column D - Equipment Category
    BARCODE: 4,         // Column E
    ASSET_SERIAL: 5,    // Column F - Asset Serial Number
    STATUS: 6,          // Column G
    OWNER: 7,           // Column H
    LOCATION: 8         // Column I
  };
  
  const formattedData = [];
  
  // Process each row (skip header row)
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // Format the row according to the new structure
    const formattedRow = [
      row[COLUMNS.CATEGORY] || '',           // A: Category
      row[COLUMNS.LOCATION] || '',           // B: Location
      row[COLUMNS.STATUS] || '',             // C: Status
      row[COLUMNS.EQUIPMENT] || '',          // D: Equip Name
      row[COLUMNS.OWNER] || '',              // E: Owner
      row[COLUMNS.UUID] || '',               // F: Equipment ID (UUID)
      row[COLUMNS.ASSET_ID] || '',           // G: Asset ID
      row[COLUMNS.ASSET_SERIAL] || '',       // H: Serial Number
      row[COLUMNS.BARCODE] || ''             // I: Barcode
    ];
    
    formattedData.push(formattedRow);
  }
  
  Logger.log(`✅ Formatted ${formattedData.length} rows for secondary export`);
  return formattedData;
} 