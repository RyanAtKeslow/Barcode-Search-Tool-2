/**
 * Update Camera's Location
 * 
 * This script fetches location data from the Asset Database and updates
 * the location column in all Camera Database sheets that contain "Status"
 * 
 * Process:
 * 1. Fetch all data from Asset Database "Barcode Database" sheet
 * 2. Identify all Camera Database sheets containing "Status"
 * 3. For each sheet, fetch barcode/serial data and update locations in bulk
 * 
 * DEBUG MODE: Set debugSheetName to target only one specific sheet for testing
 */

function updateCameraLocations() {
  Logger.log("=== Starting Camera Location Update Process ===");
  
  // DEBUG MODE: Set this to a specific sheet name to test on only one sheet
  // Leave as null/empty to process all Status sheets
  const debugSheetName = null; // Example: "Camera Status - LA" or null for all sheets
  
  if (debugSheetName) {
    Logger.log(`DEBUG MODE: Targeting only sheet: "${debugSheetName}"`);
  } else {
    Logger.log("PRODUCTION MODE: Processing all Status sheets");
  }
  
  // Step 1: Access the Asset Database and fetch all data at once
  const assetDatabaseId = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
  const cameraDatabaseId = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
  
  Logger.log("Fetching Asset Database data...");
  const assetSpreadsheet = SpreadsheetApp.openById(assetDatabaseId);
  const barcodeSheet = assetSpreadsheet.getSheetByName('Barcode Database');
  
  if (!barcodeSheet) {
    Logger.log("ERROR: 'Barcode Database' sheet not found in Asset Database");
    return;
  }
  
  // Fetch all data from Asset Database: Barcode (H), Serial (I), Location (B)
  const assetData = barcodeSheet.getRange(2, 1, barcodeSheet.getLastRow() - 1, 9).getValues();
  Logger.log(`Fetched ${assetData.length} rows from Asset Database`);
  
  // Create lookup maps for efficient matching
  const barcodeToLocation = new Map();
  const serialToLocation = new Map();
  
  for (let row of assetData) {
    const location = row[1]; // Column B
    const barcode = row[7];  // Column H
    const serial = row[8];   // Column I
    
    if (barcode && location) {
      barcodeToLocation.set(barcode.toString().trim(), location);
    }
    if (serial && location) {
      serialToLocation.set(serial.toString().trim(), location);
    }
  }
  
  Logger.log(`Created lookup maps: ${barcodeToLocation.size} barcode entries, ${serialToLocation.size} serial entries`);
  
  // Step 2: Access Camera Database and identify target sheets
  const cameraSpreadsheet = SpreadsheetApp.openById(cameraDatabaseId);
  const allSheets = cameraSpreadsheet.getSheets();
  
  let targetSheets = [];
  
  if (debugSheetName) {
    // DEBUG MODE: Find only the specified sheet
    const debugSheet = allSheets.find(sheet => sheet.getName() === debugSheetName);
    if (debugSheet) {
      targetSheets = [debugSheet];
      Logger.log(`DEBUG MODE: Found target sheet: "${debugSheetName}"`);
    } else {
      Logger.log(`ERROR: Debug sheet "${debugSheetName}" not found in Camera Database`);
      Logger.log(`Available sheets: ${allSheets.map(s => s.getName()).join(', ')}`);
      return;
    }
  } else {
    // PRODUCTION MODE: Find all Status sheets (case-insensitive)
    targetSheets = allSheets.filter(sheet => sheet.getName().toLowerCase().includes('status'));
    Logger.log(`Found ${targetSheets.length} sheets containing 'Status' (case-insensitive): ${targetSheets.map(s => s.getName()).join(', ')}`);
  }
  
  if (targetSheets.length === 0) {
    Logger.log("No target sheets found to process");
    return;
  }
  
  // Step 3: Process each target sheet
  let totalUpdated = 0;
  let totalSheetsProcessed = 0;
  
  for (let sheet of targetSheets) {
    Logger.log(`Processing sheet: ${sheet.getName()}`);
    
    try {
      const result = updateSheetLocations(sheet, barcodeToLocation, serialToLocation);
      if (result && result.updatedCount) {
        totalUpdated += result.updatedCount;
      }
      totalSheetsProcessed++;
    } catch (error) {
      Logger.log(`ERROR processing sheet ${sheet.getName()}: ${error.message}`);
    }
  }
  
  Logger.log(`=== Camera Location Update Process Complete ===`);
  Logger.log(`Total sheets processed: ${totalSheetsProcessed}`);
  Logger.log(`Total locations updated: ${totalUpdated}`);
}

/**
 * Updates locations for a single sheet in bulk
 */
function updateSheetLocations(sheet, barcodeToLocation, serialToLocation) {
  const sheetName = sheet.getName();
  
  // Fetch all data from the sheet at once
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Logger.log(`Sheet ${sheetName}: No data rows found (minimum 3 required)`);
    return;
  }
  
  const data = sheet.getRange(3, 1, lastRow - 2, 6).getValues(); // A3:F (barcode, serial, location columns)
  Logger.log(`Sheet ${sheetName}: Processing ${data.length} data rows`);
  
  // Prepare bulk update data
  const locationUpdates = [];
  let updatedCount = 0;
  let noMatchCount = 0;
  let noChangeCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const barcode = row[2]; // Column C (barcode)
    const serial = row[3];  // Column D (serial)
    const currentLocation = row[5]; // Column F (current location)
    
    let newLocation = null;
    let matchType = '';
    
    // Try to find location by barcode first, then by serial
    if (barcode && barcodeToLocation.has(barcode.toString().trim())) {
      newLocation = barcodeToLocation.get(barcode.toString().trim());
      matchType = 'barcode';
    } else if (serial && serialToLocation.has(serial.toString().trim())) {
      newLocation = serialToLocation.get(serial.toString().trim());
      matchType = 'serial';
    }
    
    if (newLocation && newLocation !== currentLocation) {
      locationUpdates.push({
        row: i + 3, // +3 because we started from row 3
        barcode: barcode,
        serial: serial,
        oldLocation: currentLocation,
        newLocation: newLocation,
        matchType: matchType
      });
      updatedCount++;
    } else if (newLocation && newLocation === currentLocation) {
      noChangeCount++;
    } else if (!newLocation) {
      noMatchCount++;
    }
  }
  
  Logger.log(`Sheet ${sheetName}: ${updatedCount} locations to update, ${noChangeCount} no changes needed, ${noMatchCount} no matches found`);
  
  // Perform bulk update if there are changes
  if (locationUpdates.length > 0) {
    // Log detailed information about what will be updated
    Logger.log(`Sheet ${sheetName}: Preparing to update ${updatedCount} rows:`);
    for (let update of locationUpdates) {
      Logger.log(`  Row ${update.row}: ${update.oldLocation} → ${update.newLocation} (matched by ${update.matchType})`);
      if (update.barcode) Logger.log(`    Barcode: ${update.barcode}`);
      if (update.serial) Logger.log(`    Serial: ${update.serial}`);
    }
    
    // Group updates by location value to minimize range operations
    const locationGroups = new Map();
    for (let update of locationUpdates) {
      if (!locationGroups.has(update.newLocation)) {
        locationGroups.set(update.newLocation, []);
      }
      locationGroups.get(update.newLocation).push(update.row);
    }
    
    // Update each location value in bulk
    for (let [location, rows] of locationGroups) {
      if (rows.length === 1) {
        // Single row update
        sheet.getRange(rows[0], 6).setValue(location); // Column F
      } else {
        // Multiple rows with same location - update in one operation
        const range = sheet.getRange(rows[0], 6, rows.length, 1); // Column F, multiple rows
        range.setValue(location);
      }
    }
    
    Logger.log(`Sheet ${sheetName}: Successfully updated ${updatedCount} locations in bulk`);
  } else {
    Logger.log(`Sheet ${sheetName}: No location updates needed`);
  }
  
  // Return summary for this sheet
  return {
    sheetName: sheetName,
    updatedCount: updatedCount,
    noChangeCount: noChangeCount,
    noMatchCount: noMatchCount,
    totalRows: data.length
  };
}

/**
 * Test function to run the update process
 */
function testUpdateCameraLocations() {
  Logger.log("Running test update...");
  updateCameraLocations();
}

/**
 * Debug function to test on a specific sheet only
 * @param {string} sheetName - The exact name of the sheet to test on
 */
function debugUpdateCameraLocations(sheetName) {
  if (!sheetName) {
    Logger.log("ERROR: Please provide a sheet name for debug mode");
    Logger.log("Usage: debugUpdateCameraLocations('Sheet Name Here')");
    return;
  }
  
  Logger.log(`=== DEBUG MODE: Testing on sheet "${sheetName}" ===`);
  
  // Temporarily set debug mode and run
  const originalFunction = updateCameraLocations;
  
  // Create a modified version for this debug run
  const debugFunction = function() {
    Logger.log("=== Starting Camera Location Update Process (DEBUG MODE) ===");
    
    // Set debug mode for this run
    const debugSheetName = sheetName;
    Logger.log(`DEBUG MODE: Targeting only sheet: "${debugSheetName}"`);
    
    // Step 1: Access the Asset Database and fetch all data at once
    const assetDatabaseId = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
    const cameraDatabaseId = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
    
    Logger.log("Fetching Asset Database data...");
    const assetSpreadsheet = SpreadsheetApp.openById(assetDatabaseId);
    const barcodeSheet = assetSpreadsheet.getSheetByName('Barcode Database');
    
    if (!barcodeSheet) {
      Logger.log("ERROR: 'Barcode Database' sheet not found in Asset Database");
      return;
    }
    
    // Fetch all data from Asset Database: Barcode (H), Serial (I), Location (B)
    const assetData = barcodeSheet.getRange(2, 1, barcodeSheet.getLastRow() - 1, 9).getValues();
    Logger.log(`Fetched ${assetData.length} rows from Asset Database`);
    
    // Create lookup maps for efficient matching
    const barcodeToLocation = new Map();
    const serialToLocation = new Map();
    
    for (let row of assetData) {
      const location = row[1]; // Column B
      const barcode = row[7];  // Column H
      const serial = row[8];   // Column I
      
      if (barcode && location) {
        barcodeToLocation.set(barcode.toString().trim(), location);
      }
      if (serial && location) {
        serialToLocation.set(serial.toString().trim(), location);
      }
    }
    
    Logger.log(`Created lookup maps: ${barcodeToLocation.size} barcode entries, ${serialToLocation.size} serial entries`);
    
    // Step 2: Access Camera Database and find the specific sheet
    const cameraSpreadsheet = SpreadsheetApp.openById(cameraDatabaseId);
    const allSheets = cameraSpreadsheet.getSheets();
    
    const debugSheet = allSheets.find(sheet => sheet.getName() === debugSheetName);
    if (debugSheet) {
      Logger.log(`DEBUG MODE: Found target sheet: "${debugSheetName}"`);
    } else {
      Logger.log(`ERROR: Debug sheet "${debugSheetName}" not found in Camera Database`);
      Logger.log(`Available sheets: ${allSheets.map(s => s.getName()).join(', ')}`);
      return;
    }
    
    // Step 3: Process the debug sheet
    Logger.log(`Processing sheet: ${debugSheet.getName()}`);
    
    try {
      updateSheetLocations(debugSheet, barcodeToLocation, serialToLocation);
    } catch (error) {
      Logger.log(`ERROR processing sheet ${debugSheet.getName()}: ${error.message}`);
    }
    
    Logger.log("=== Camera Location Update Process Complete (DEBUG MODE) ===");
  };
  
  // Run the debug function
  debugFunction();
} 