/**
 * Count Barcodes For Analytics - Barcode Counting and Analytics Script
 * 
 * This script counts barcodes across multiple bin sheets and updates the Analytics
 * sheet with comprehensive barcode statistics for inventory tracking.
 * 
 * Step-by-step process:
 * 1. Defines target sheets (ER Aisles 1-15, Service Department, Toolchests, etc.)
 * 2. Reads Barcode Dictionary data from columns D3:G
 * 3. Precomputes pipe counts for each item in the dictionary
 * 4. Calculates total dictionary barcode count (pipes + 1 per row)
 * 5. Processes each target sheet individually
 * 6. Extracts unique item names from column B2:B of each sheet
 * 7. Matches item names against dictionary pipe counts
 * 8. Sums total barcodes per sheet based on dictionary matches
 * 9. Writes results to Analytics sheet (columns V and W)
 * 10. Updates total dictionary count in cell C2
 * 
 * Target Sheets:
 * - RnD, ER Aisles 1-15
 * - Service Department, Battery Room, Filter Room
 * - S-Toolchest A-Z (65-90)
 * - Purchasing Mezzanine, Projector Room
 * - Consignment Rooms, Inventory Control
 * 
 * Data Processing:
 * - Pipe counting: Counts | characters + 1 per row for barcode totals
 * - Item matching: Matches sheet item names to dictionary entries
 * - Batch processing: Handles large datasets efficiently
 * - Progress logging: Tracks processing every 1000 rows
 * 
 * Analytics Output:
 * - Column V: Sheet names
 * - Column W: Barcode counts per sheet
 * - Cell C2: Total dictionary barcode count
 * - Row-by-row tracking for each processed sheet
 * 
 * Features:
 * - Comprehensive sheet coverage (20+ sheets)
 * - Efficient batch processing
 * - Detailed progress logging
 * - Accurate pipe-delimited barcode counting
 * - Real-time analytics updates
 */
function countBarcodesForAnalytics() {
  Logger.log("Starting countBarcodesForAnalytics function");
  
  // Get the active spreadsheet and the active sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Active spreadsheet obtained");
  const barcodeSheet = ss.getSheetByName("Barcode Dictionary");
  Logger.log("Barcode Dictionary sheet obtained");

  // Target sheet names
  const sheetNames = ["RnD"];
  for (let i = 1; i <= 15; i++) sheetNames.push("ER Aisle " + i);
  sheetNames.push("Service Department");
  for (let i = 65; i <= 90; i++) sheetNames.push("S-Toolchest " + String.fromCharCode(i));
  sheetNames.push("Battery Room", "Filter Room", "Purchasing Mezzanine", "Projector Room", "Consignment Rooms", "Inventory Control");

  const analyticsSheet = ss.getSheetByName("Analytics");
  Logger.log("Analytics sheet obtained");

  // Initialize row index for writing to Analytics starting at row 2
  let analyticsRowIndex = 2;

  // Get all rows in column D and G of the Barcode Dictionary
  const dictionaryValues = barcodeSheet.getRange("D3:G").getValues();
  const pipeCountMap = new Map();

  // Precompute pipe counts for the specific item in the dictionary
  dictionaryValues.forEach((row, index) => {
    const itemName = row[0].trim();
    const columnGValue = String(row[3]); // Column G is the 4th column in the range D3:G
    const pipeCount = (columnGValue.match(/\|/g) || []).length + 1;
    Logger.log(`Processing ${itemName} at row ${index + 3}`);
    Logger.log(`Column G value: ${columnGValue}`);
    Logger.log(`Observed pipe count: ${(columnGValue.match(/\|/g) || []).length}`);
    Logger.log(`Final pipe count (including +1): ${pipeCount}`);
    if (pipeCountMap.has(itemName)) {
      pipeCountMap.set(itemName, pipeCountMap.get(itemName) + pipeCount);
    } else {
      pipeCountMap.set(itemName, pipeCount);
    }
  });

  Logger.log("Starting to count total pipes in column G")
  // Count total pipes in column G
  let totalDictionaryPipeCount = 0;
  let rowCount = 0;
  dictionaryValues.forEach((row, index) => {
    const columnGValue = String(row[3]); // Column G is the 4th column in the range D3:G
    const pipeCount = (columnGValue.match(/\|/g) || []).length + 1;
    totalDictionaryPipeCount += pipeCount;
    rowCount++;
    if (index % 1000 === 0) {
      Logger.log(`Processed ${rowCount} rows. Current total: ${totalDictionaryPipeCount}`);
    }
  });
  Logger.log(`Final row count: ${rowCount}`);
  Logger.log(`Total pipe count in Barcode Dictionary G3:G (including +1 per row): ${totalDictionaryPipeCount}`);

  // Iterate over each target sheet
  sheetNames.forEach(sheetName => {
    const activeSheet = ss.getSheetByName(sheetName);
    if (!activeSheet) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return;
    }
    Logger.log(`Processing sheet: ${sheetName}`);

    // Get all unique item names from column B2:B
    const itemNames = [...new Set(activeSheet.getRange("B2:B").getValues().flat().filter(name => name !== "").map(name => name.trim()))];
    Logger.log(`Unique item names obtained from ${sheetName}: ${itemNames.length} unique items found`);
    itemNames.forEach(name => {
      Logger.log(`Item name: ${name}`);
      if (pipeCountMap.has(name)) {
        Logger.log(`Pipe count for ${name}: ${pipeCountMap.get(name)}`);
      } else {
        Logger.log(`No pipe count found for ${name}`);
      }
    });

    // Initialize a counter for the total pipe count
    let totalPipeCount = 0;

    // Iterate over each item name
    itemNames.forEach(itemName => {
      if (pipeCountMap.has(itemName)) {
        totalPipeCount += pipeCountMap.get(itemName);
      }
    });

    // Append sheet name and pipe count to Analytics
    analyticsSheet.getRange(analyticsRowIndex, 22).setValue(sheetName); // Column V
    analyticsSheet.getRange(analyticsRowIndex, 23).setValue(totalPipeCount); // Column W
    Logger.log(`Appended to Analytics: ${sheetName}, ${totalPipeCount}`);
    analyticsRowIndex++;
  });

  // Write the total dictionary pipe count to cell C2 of Analytics
  Logger.log(`Writing total pipe count ${totalDictionaryPipeCount} to C2`);
  analyticsSheet.getRange("C2").setValue(totalDictionaryPipeCount);
  Logger.log(`Total dictionary pipe count written to C2: ${totalDictionaryPipeCount}`);

  Logger.log("countBarcodesForAnalytics function completed");
} 