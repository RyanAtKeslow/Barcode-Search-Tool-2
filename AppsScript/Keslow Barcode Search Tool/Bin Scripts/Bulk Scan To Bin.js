/**
 * Bulk Scan To Bin - Bulk Data Processing Script
 * 
 * This script processes large batches of barcode data from the "Bulk Scan" sheet
 * and distributes them to appropriate bin sheets based on naming conventions.
 * 
 * Step-by-step process:
 * 1. Reads all data from "Bulk Scan" sheet (barcode, equipment, proposed bin, rectified status)
 * 2. Filters out already processed rows and rows with missing data
 * 3. Groups data by target sheet based on bin naming conventions:
 *    - Number bins (1-10) → "ER Aisle X" sheets
 *    - Letter bins (S,B,F,M,P,R,C,Q) → specific department sheets
 * 4. Processes each sheet group in bulk operations
 * 5. Checks for duplicate bin+equipment combinations
 * 6. Inserts new rows below target bins with equipment data
 * 7. Updates rectified status in bulk for processed items
 * 8. Displays comprehensive summary with success/error counts
 * 
 * Bin Naming Conventions:
 * - S → Service Department
 * - B → Battery Room  
 * - F → Filter Room
 * - M → Mezzanine
 * - P → Projector Room
 * - R → RnD
 * - C → Consignment Rooms
 * - Q → Inventory Control
 * - 1-10 → ER Aisle 1-10
 * 
 * Features:
 * - Duplicate detection and handling
 * - Bulk operations for performance
 * - Comprehensive error reporting
 * - Progress tracking and validation
 */
function bulkScanToBin() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the Bulk Scan sheet
  const bulkScanSheet = spreadsheet.getSheetByName("Bulk Scan");
  if (!bulkScanSheet) {
    ui.alert("Error: 'Bulk Scan' sheet not found.");
    return;
  }
  
  // Get all data from Bulk Scan sheet (excluding header)
  const lastRow = bulkScanSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert("No data found in Bulk Scan sheet.");
    return;
  }
  
  const data = bulkScanSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  // Filter out rows that are already rectified or have empty cells
  const pendingData = data.filter(row => {
    const [barcode, equipment, proposedBin, rectified] = row;
    // Skip if already rectified or if any required field is empty
    return !rectified && 
           barcode && barcode.toString().trim() !== '' && 
           equipment && equipment.toString().trim() !== '' && 
           proposedBin && proposedBin.toString().trim() !== '';
  });
  
  if (pendingData.length === 0) {
    const rectifiedRows = data.filter(row => row[3]).length;
    
    ui.alert("No data to process", 
      `Already rectified: ${rectifiedRows} rows\n\n` +
      `All rows are either completed or missing required data.`, 
      ui.ButtonSet.OK);
    return;
  }
  
  // Group data by target sheet for bulk operations
  const sheetGroups = groupDataBySheet(pendingData);
  
  // Filter out items with invalid sheet names
  const validSheetGroups = {};
  const invalidItems = [];
  
  Object.keys(sheetGroups).forEach(sheetName => {
    if (sheetName && sheetName !== 'null') {
      validSheetGroups[sheetName] = sheetGroups[sheetName];
    } else {
      invalidItems.push(...sheetGroups[sheetName]);
    }
  });
  
  if (invalidItems.length > 0) {
    console.warn(`Found ${invalidItems.length} items with invalid bin names:`, invalidItems.map(item => item.proposedBin));
  }
  
  // Process each sheet group
  const results = processSheetGroups(spreadsheet, validSheetGroups);
  
  // Add invalid items to results for reporting
  if (invalidItems.length > 0) {
    results.invalidItems = invalidItems;
  }
  
  // Update rectified column in bulk (only for successfully processed items)
  updateRectifiedStatus(bulkScanSheet, results.success);
  
  // Show summary
  showResultsSummary(results);
}

/**
 * Groups data by target sheet based on bin naming conventions
 */
function groupDataBySheet(data) {
  const sheetGroups = {};
  
  data.forEach(row => {
    const [barcode, equipment, proposedBin, rectified] = row;
    const targetSheet = getTargetSheetName(proposedBin);
    
    if (!sheetGroups[targetSheet]) {
      sheetGroups[targetSheet] = [];
    }
    
    sheetGroups[targetSheet].push({
      barcode: barcode,
      equipment: equipment,
      proposedBin: proposedBin,
      originalRow: row
    });
  });
  
  return sheetGroups;
}

/**
 * Determines target sheet name based on bin naming convention
 */
function getTargetSheetName(binName) {
  if (!binName || typeof binName !== 'string') return null;
  
  const trimmedBin = binName.trim();
  if (trimmedBin.length === 0) return null;
  
  const firstChar = trimmedBin.charAt(0);
  
  // Check if bin starts with a number (1-10)
  if (/^[1-9]$/.test(firstChar)) {
    return `ER Aisle ${firstChar}`;
  }
  
  // Check for double-digit numbers (10)
  if (trimmedBin.startsWith('10')) {
    return 'ER Aisle 10';
  }
  
  // Check letter-based naming convention
  const letterMap = {
    'S': 'Service Department',
    'B': 'Battery Room',
    'F': 'Filter Room',
    'M': 'Mezzanine',
    'P': 'Projector Room',
    'R': 'RnD',
    'C': 'Consignment Rooms',
    'Q': 'Inventory Control'
  };
  
  return letterMap[firstChar] || null;
}

/**
 * Processes data for each sheet group in bulk operations
 */
function processSheetGroups(spreadsheet, sheetGroups) {
  const results = {
    success: [],
    errors: [],
    summary: {}
  };
  
  Object.keys(sheetGroups).forEach(sheetName => {
    try {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        results.errors.push(`Sheet '${sheetName}' not found`);
        return;
      }
      
      const groupData = sheetGroups[sheetName];
      const groupResult = processSheetGroup(sheet, groupData);
      
      results.success.push(...groupResult.success);
      results.errors.push(...groupResult.errors);
      
      // Update summary (count both new insertions and duplicates)
      if (!results.summary[sheetName]) {
        results.summary[sheetName] = 0;
      }
      results.summary[sheetName] += groupResult.success.length;
      
    } catch (error) {
      results.errors.push(`Error processing sheet '${sheetName}': ${error.message}`);
    }
  });
  
  return results;
}

/**
 * Processes a group of data for a specific sheet
 */
function processSheetGroup(sheet, groupData) {
  const results = {
    success: [],
    errors: []
  };
  
  // Get all data from the target sheet (columns A and B)
  const lastRow = sheet.getLastRow();
  const sheetData = sheet.getRange(1, 1, lastRow, 2).getValues();
  
  // Create a map of bin positions for quick lookup
  const binPositions = {};
  // Create a set of existing bin+equipment combinations to check for duplicates
  const existingCombinations = new Set();
  
  sheetData.forEach((row, index) => {
    const [bin, equipment] = row;
    if (bin && typeof bin === 'string') {
      binPositions[bin.trim()] = index + 1; // +1 for 1-based indexing
      
      // Add existing combinations to the set for duplicate checking
      if (equipment && typeof equipment === 'string') {
        existingCombinations.add(`${bin.trim()}|${equipment.trim()}`);
      }
    }
  });
  
  // Process each item in the group
  groupData.forEach(item => {
    try {
      const targetRow = binPositions[item.proposedBin];
      
      if (!targetRow) {
        results.errors.push(`Bin '${item.proposedBin}' not found in sheet '${sheet.getName()}'`);
        return;
      }
      
      // Check if this bin+equipment combination already exists
      const combinationKey = `${item.proposedBin}|${item.equipment}`;
      if (existingCombinations.has(combinationKey)) {
        // Duplicate found - mark as success since it's already in the database
        results.success.push({
          barcode: item.barcode,
          bin: item.proposedBin,
          sheet: sheet.getName(),
          row: null, // No new row inserted
          isDuplicate: true
        });
        return;
      }
      
      // Insert new row below the target bin
      sheet.insertRowAfter(targetRow);
      
      // Insert data in the new row
      const insertRow = targetRow + 1;
      sheet.getRange(insertRow, 1).setValue(item.proposedBin);
      sheet.getRange(insertRow, 2).setValue(item.equipment);
      
      // Add to existing combinations set to prevent duplicates within the same batch
      existingCombinations.add(combinationKey);
      
      results.success.push({
        barcode: item.barcode,
        bin: item.proposedBin,
        sheet: sheet.getName(),
        row: insertRow
      });
      
    } catch (error) {
      results.errors.push(`Error processing item ${item.barcode}: ${error.message}`);
    }
  });
  
  return results;
}

/**
 * Updates the rectified column for processed rows
 */
function updateRectifiedStatus(bulkScanSheet, processedData) {
  if (processedData.length === 0) return;
  
  // Find the row numbers for processed data
  const lastRow = bulkScanSheet.getLastRow();
  const allData = bulkScanSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  const rowsToUpdate = [];
  processedData.forEach(processedItem => {
    // processedItem has: { barcode, bin, sheet, row }
    // We need to find the original row in Bulk Scan sheet by matching barcode
    const rowIndex = allData.findIndex(row => 
      row[0] === processedItem.barcode // Match barcode in column A
    );
    
    if (rowIndex !== -1) {
      rowsToUpdate.push(rowIndex + 2); // +2 for header row and 0-based index
    }
  });
  
  // Update rectified column in bulk
  rowsToUpdate.forEach(rowNum => {
    bulkScanSheet.getRange(rowNum, 4).setValue(true);
  });
}

/**
 * Shows a summary of the processing results
 */
function showResultsSummary(results) {
  const ui = SpreadsheetApp.getUi();
  
  let message = "Bulk Scan To Bin Processing Complete!\n\n";
  
  // Summary by sheet
  message += "Items processed by sheet:\n";
  Object.keys(results.summary).forEach(sheetName => {
    message += `• ${sheetName}: ${results.summary[sheetName]} items\n`;
  });
  
  // Separate duplicates from new insertions
  const newInsertions = results.success.filter(item => !item.isDuplicate);
  const duplicates = results.success.filter(item => item.isDuplicate);
  
  message += `\nTotal successful: ${results.success.length}`;
  if (newInsertions.length > 0) {
    message += ` (${newInsertions.length} new insertions`;
    if (duplicates.length > 0) {
      message += `, ${duplicates.length} duplicates found`;
    }
    message += `)`;
  }
  
  if (results.invalidItems && results.invalidItems.length > 0) {
    message += `\nTotal skipped (invalid bin names): ${results.invalidItems.length}`;
    message += "\n\nSkipped items:\n";
    results.invalidItems.slice(0, 5).forEach(item => {
      message += `• ${item.barcode} → ${item.proposedBin}\n`;
    });
    
    if (results.invalidItems.length > 5) {
      message += `... and ${results.invalidItems.length - 5} more skipped items.`;
    }
  }
  
  if (results.errors.length > 0) {
    message += `\nTotal errors: ${results.errors.length}`;
    message += "\n\nErrors:\n";
    results.errors.slice(0, 5).forEach(error => {
      message += `• ${error}\n`;
    });
    
    if (results.errors.length > 5) {
      message += `... and ${results.errors.length - 5} more errors.`;
    }
  }
  
  ui.alert("Processing Complete", message, ui.ButtonSet.OK);
  
  // Log results to console for debugging
  console.log("Bulk Scan Results:", results);
}

/**
 * Test function to run the bulk scan process
 */
function testBulkScanToBin() {
  try {
    bulkScanToBin();
  } catch (error) {
    console.error("Error in bulk scan process:", error);
    SpreadsheetApp.getUi().alert("Error", `An error occurred: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
