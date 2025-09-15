/**
 * Refresh Analytics Weekly - Weekly Analytics Data Collection Script
 * 
 * This script collects weekly analytics data and performs barcode dictionary
 * validation checks for the Lost & Found system.
 * 
 * Step-by-step process:
 * 1. Opens the Analytics sheet and finds the first empty row in column P
 * 2. Collects source values from B3:B6 (analytics metrics)
 * 3. Writes values to columns P, Q, R, S in the empty row
 * 4. Formats column R as text (ratio) and column S as percentage
 * 5. Adds today's date in column O
 * 6. Creates difference formulas in column T for trend analysis
 * 7. Runs barcode counting function for additional analytics
 * 8. Performs Lost & Found barcode validation against dictionary
 * 
 * Analytics Data:
 * - Column P: Primary metric from B3
 * - Column Q: Secondary metric from B4  
 * - Column R: Ratio from B5 (formatted as text)
 * - Column S: Percentage from B6 (formatted as percentage)
 * - Column O: Date stamp
 * - Column T: Difference calculations for trend analysis
 * 
 * Lost & Found Validation:
 * - Compares Lost & Found barcodes against Barcode Dictionary
 * - Updates column J to TRUE for barcodes with "active" status
 * - Handles pipe-delimited barcode formats
 * - Provides detailed logging of validation results
 * 
 * Features:
 * - Automatic data collection and formatting
 * - Trend analysis with difference calculations
 * - Barcode validation and status updates
 * - Comprehensive logging and error handling
 * - Integration with barcode counting analytics
 */
function refreshAnalyticsWeekly() {
  // Get the active spreadsheet and the Analytics sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Analytics");
  
  // Find the first empty row in column P
  const lastRow = sheet.getLastRow();
  let firstEmptyRow = 1;
  
  for (let i = 1; i <= lastRow + 1; i++) {
    if (sheet.getRange(i, 16).getValue() === "") { // Column P is the 16th column
      firstEmptyRow = i;
      break;
    }
  }
  
  // Get the values from B3:B6
  const sourceValues = sheet.getRange("B3:B6").getValues();
  
  // Log the source values for debugging
  Logger.log("Source values from B3:B6:");
  sourceValues.forEach((value, index) => {
    Logger.log(`B${index + 3}: ${value[0]}`);
  });
  
  // Write the values to PQRS columns in the found row
  for (let i = 0; i < sourceValues.length; i++) {
    const value = sourceValues[i][0];
    const targetRange = sheet.getRange(firstEmptyRow, 16 + i);
    
    if (i === 2) { // For the ratio (column R)
      // Force the ratio to be treated as text
      targetRange.setNumberFormat('@');
      targetRange.setValue("'" + value); // Add a single quote to force text
    } else if (i === 3) { // For the percentage (column S)
      // Set percentage format
      targetRange.setNumberFormat('0.00%');
      targetRange.setValue(value);
    } else {
      // For other columns, write as is
      targetRange.setValue(value);
    }
    
    Logger.log(`Writing to column ${String.fromCharCode(80 + i)}${firstEmptyRow}: ${value}`);
  }
  
  // Write today's date in column O (one column to the left of P)
  sheet.getRange(firstEmptyRow, 15).setValue(new Date());
  
  // Write the difference formula in column T, one row above the found row
  if (firstEmptyRow > 1) {
    const formula = `=P${firstEmptyRow}-P${firstEmptyRow-1}`;
    sheet.getRange(firstEmptyRow - 1, 20).setFormula(formula); // Column T is the 20th column
  }
  // Write the formula "=B3-P(firstEmptyRow)" in column T one row below the first empty row

  const formula = `=B3-P${firstEmptyRow}`;
  sheet.getRange(firstEmptyRow, 20).setFormula(formula);


  // Run the barcode counting function
  countBarcodesForAnalytics();
}

function checkLostAndFoundAgainstDictionary() {
  Logger.log("Starting Lost & Found barcode dictionary check");
  
  // Get the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get both sheets
  const lostFoundSheet = ss.getSheetByName("Lost & Found");
  const dictionarySheet = ss.getSheetByName("Barcode Dictionary");
  
  if (!lostFoundSheet) {
    Logger.log("Error: Lost & Found sheet not found");
    return;
  }
  
  if (!dictionarySheet) {
    Logger.log("Error: Barcode Dictionary sheet not found");
    return;
  }
  
  // Get data from Lost & Found sheet (columns A and C)
  const lostFoundLastRow = lostFoundSheet.getLastRow();
  if (lostFoundLastRow < 2) {
    Logger.log("No data found in Lost & Found sheet");
    return;
  }
  
  const lostFoundData = lostFoundSheet.getRange(2, 1, lostFoundLastRow - 1, 3).getValues(); // A, B, C columns
  Logger.log(`Found ${lostFoundData.length} rows in Lost & Found sheet`);
  
  // Get data from Barcode Dictionary sheet (columns G and C)
  const dictionaryLastRow = dictionarySheet.getLastRow();
  if (dictionaryLastRow < 2) {
    Logger.log("No data found in Barcode Dictionary sheet");
    return;
  }
  
  const dictionaryData = dictionarySheet.getRange(2, 1, dictionaryLastRow - 1, 7).getValues(); // Get columns A through G
  Logger.log(`Found ${dictionaryData.length} rows in Barcode Dictionary sheet`);
  
  // Create a lookup map for dictionary data
  const dictionaryMap = new Map();
  dictionaryData.forEach((row, index) => {
    const barcode = row[6]; // Column G (index 6)
    const status = row[2];  // Column C (index 2)
    
    if (barcode && status) {
      // Clean the barcode by removing pipes
      const cleanBarcode = barcode.toString().replace(/^\|?([^|]+)\|?$/, '$1');
      dictionaryMap.set(cleanBarcode, status.toString());
    }
  });
  
  Logger.log(`Created dictionary map with ${dictionaryMap.size} entries`);
  
  // Process each row in Lost & Found
  let updatedCount = 0;
  lostFoundData.forEach((row, index) => {
    const barcode = row[0]; // Column A
    const lostFoundStatus = row[2]; // Column C
    
    if (barcode) {
      const cleanBarcode = barcode.toString().trim();
      const dictionaryStatus = dictionaryMap.get(cleanBarcode);
      
      if (dictionaryStatus && (dictionaryStatus.toLowerCase() === 'active')) {
        // Set column J (index 10) to TRUE
        const rowNumber = index + 2; // +2 because we started from row 2
        lostFoundSheet.getRange(rowNumber, 10).setValue(true);
        updatedCount++;
        
        Logger.log(`Updated row ${rowNumber}: Barcode ${cleanBarcode} - Dictionary status: ${dictionaryStatus}, Lost & Found status: ${lostFoundStatus}`);
      }
    }
  });
  
  Logger.log(`Completed Lost & Found check. Updated ${updatedCount} rows.`);
} 