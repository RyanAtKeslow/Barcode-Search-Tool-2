/**
 * Export To Secondary Database - External Database Export Script
 * 
 * This script exports processed barcode data to an external secondary database
 * with proper formatting and column mapping for asset management.
 * 
 * Step-by-step process:
 * 1. Connects to target spreadsheet (1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM)
 * 2. Creates or accesses 'Barcode Database' sheet
 * 3. Formats data using formatDataForSecondaryExport helper
 * 4. Clears target sheet and removes excess columns
 * 5. Adds completion timestamp and record count
 * 6. Writes formatted headers and data
 * 7. Sets frozen rows for better navigation
 * 8. Returns success/failure status with details
 * 
 * Data Formatting:
 * - Headers: Category, Location, Status, Equip Name, Owner, Equipment ID, Asset ID, Serial Number, Barcode
 * - Column mapping: Reorganizes data from source format to target format
 * - Data validation: Handles empty values and missing data gracefully
 * 
 * Target Database:
 * - Spreadsheet ID: 1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM
 * - Sheet name: Barcode Database
 * - Format: Standardized asset management format
 * 
 * Column Mapping:
 * - Source columns: Asset ID, UUID, Equipment, Category, Barcode, Asset Serial, Status, Owner, Location
 * - Target columns: Category, Location, Status, Equip Name, Owner, Equipment ID, Asset ID, Serial Number, Barcode
 * 
 * Features:
 * - External database integration
 * - Data format transformation
 * - Column cleanup and optimization
 * - Timestamp and record tracking
 * - Comprehensive error handling
 * - Success/failure reporting
 */
function exportToSecondaryDatabase(processedData, summaryStats) {
  Logger.log("🔄 Starting secondary database export...");
  
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