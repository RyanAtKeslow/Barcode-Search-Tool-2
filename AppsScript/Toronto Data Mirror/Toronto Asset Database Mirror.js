/**
 * Toronto Asset Database Mirror
 * 
 * This script ingests data from the Toronto Barcode Database spreadsheet
 * and remaps the columns to match the Toronto Schema format.
 * 
 * Source Schema (Barcode Database):
 * A: Category, B: Asset Location, C: Asset Status, D: Asset, E: Owner, 
 * F: Equipment ID, G: Asset ID, H: Serial Number, I: Barcode
 * 
 * Target Schema (Toronto Schema):
 * A: Asset, B: Barcode #, C: Serial #, D: Owner, E: Asset Code, 
 * F: Asset Location, G: Asset Status, H: Category
 */

function mirrorTorontoAssetDatabase() {
  try {
    // Source spreadsheet ID and sheet name
    const sourceSpreadsheetId = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
    const sourceSheetName = 'Barcode Database';
    
    // Get source spreadsheet and sheet
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    
    if (!sourceSheet) {
      throw new Error(`Sheet "${sourceSheetName}" not found in source spreadsheet`);
    }
    
    // Get all data from source sheet (columns A:I)
    const sourceData = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), 9).getValues();
    
    if (sourceData.length <= 1) {
      throw new Error('No data found in source sheet (only header row exists)');
    }
    
    // Extract data rows (skip header row)
    const dataRows = sourceData.slice(1);
    
    // Create target data with remapped columns
    const targetData = createTargetData(dataRows);
    
    // Use the same spreadsheet, target Sheet4
    const targetSheet = sourceSpreadsheet.getSheetByName('Sheet4');
    
    if (!targetSheet) {
      throw new Error('Sheet "Sheet4" not found in source spreadsheet');
    }
    
    // Clear existing data and write new data
    targetSheet.clear();
    targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
    
    // Format the sheet
    formatTargetSheet(targetSheet, targetData.length);
    
    Logger.log(`Successfully mirrored ${dataRows.length} rows to Toronto Asset Database`);
    
    return {
      success: true,
      rowsProcessed: dataRows.length,
      targetSpreadsheetUrl: sourceSpreadsheet.getUrl()
    };
    
  } catch (error) {
    Logger.log(`Error mirroring Toronto Asset Database: ${error.message}`);
    throw error;
  }
}



/**
 * Creates target data with remapped columns according to Toronto Schema
 */
function createTargetData(sourceDataRows) {
  // Toronto Schema headers
  const targetHeaders = [
    'Asset',           // D: Asset
    'Barcode #',       // I: Barcode  
    'Serial #',        // H: Serial Number
    'Owner',           // E: Owner
    'Asset Code',      // G: Asset ID
    'Asset Location',  // B: Asset Location
    'Asset Status',    // C: Asset Status
    'Category'         // A: Category
  ];
  
  const targetData = [targetHeaders];
  
  // Remap each data row
  for (const sourceRow of sourceDataRows) {
    const targetRow = [
      sourceRow[3],  // D: Asset
      sourceRow[8],  // I: Barcode
      sourceRow[7],  // H: Serial Number
      sourceRow[4],  // E: Owner
      sourceRow[6],  // G: Asset ID
      sourceRow[1],  // B: Asset Location
      sourceRow[2],  // C: Asset Status
      sourceRow[0]   // A: Category
    ];
    
    targetData.push(targetRow);
  }
  
  return targetData;
}

/**
 * Creates or gets the target spreadsheet
 */
function createOrGetTargetSpreadsheet() {
  const targetName = 'Toronto Asset Database Mirror';
  
  // Try to find existing spreadsheet
  const files = DriveApp.getFilesByName(targetName);
  
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }
  
  // Create new spreadsheet if none exists
  const newSpreadsheet = SpreadsheetApp.create(targetName);
  
  // Move to the same folder as the source spreadsheet (if possible)
  try {
    const sourceFile = DriveApp.getFileById('1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM');
    const sourceFolder = sourceFile.getParents().next();
    sourceFolder.addFile(newSpreadsheet);
    DriveApp.getRootFolder().removeFile(newSpreadsheet);
  } catch (e) {
    Logger.log('Could not move target spreadsheet to source folder: ' + e.message);
  }
  
  return newSpreadsheet;
}

/**
 * Creates or gets the target sheet within the spreadsheet
 */
function createOrGetTargetSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Formats the target sheet for better readability
 */
function formatTargetSheet(sheet, dataRowCount) {
  if (dataRowCount === 0) return;
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Auto-resize columns
  for (let i = 1; i <= 8; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Add borders to all data
  const dataRange = sheet.getRange(1, 1, dataRowCount, 8);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set alternating row colors for better readability
  if (dataRowCount > 1) {
    for (let row = 2; row <= dataRowCount; row++) {
      if (row % 2 === 0) {
        sheet.getRange(row, 1, 1, 8).setBackground('#f8f9fa');
      }
    }
  }
}

/**
 * Test function to run the mirror operation
 */
function testTorontoAssetDatabaseMirror() {
  try {
    const result = mirrorTorontoAssetDatabase();
    Logger.log('Test completed successfully:');
    Logger.log(`- Rows processed: ${result.rowsProcessed}`);
    Logger.log(`- Target spreadsheet: ${result.targetSpreadsheetUrl}`);
  } catch (error) {
    Logger.log(`Test failed: ${error.message}`);
  }
}

/**
 * Manual trigger function for running the mirror operation
 */
function runTorontoAssetDatabaseMirror() {
  return mirrorTorontoAssetDatabase();
}
