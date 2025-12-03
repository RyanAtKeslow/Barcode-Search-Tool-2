/**
 * Measure Sheet Performance - Performance Analysis Script
 * 
 * This script measures the load time and performance of all sheets in the workbook
 * to identify bottlenecks and optimize formula recalculation times.
 * 
 * Step-by-step process:
 * 1. Opens the target spreadsheet by ID
 * 2. Gets all sheets in the workbook
 * 3. For each sheet, measures:
 *    - Time to get sheet metadata (last row, last column)
 *    - Time to get values from the used range
 *    - Time to get formulas from the used range
 *    - Time to get formats from the used range
 *    - Total sheet load time
 * 4. Logs detailed results to Logger
 * 5. Optionally writes results to a "Performance Report" sheet
 * 
 * Performance Metrics:
 * - Sheet name
 * - Last row and column (data size)
 * - Values retrieval time
 * - Formulas retrieval time
 * - Formats retrieval time
 * - Total load time
 * - Estimated cell count
 * 
 * Features:
 * - Comprehensive performance testing across all sheets
 * - Multiple operation measurements
 * - Detailed logging for analysis
 * - Optional results sheet for easy viewing
 * - Handles empty sheets gracefully
 */

function measurePerformance() {
  // Spreadsheet ID from the Google Sheets URL
  const SPREADSHEET_ID = '1q4YqY_vLHsmnN6PDihGAo2seCc0Tugo_5x9QDTF1cek';
  
  Logger.log("=== Starting Performance Measurement ===");
  Logger.log("Spreadsheet ID: " + SPREADSHEET_ID);
  
  // Open the spreadsheet
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Spreadsheet opened successfully");
  } catch (e) {
    Logger.log("ERROR: Could not open spreadsheet. " + e.toString());
    return;
  }
  
  // Get all sheets
  var sheets = ss.getSheets();
  Logger.log("Total sheets found: " + sheets.length);
  Logger.log("");
  
  // Array to store results for potential sheet output
  var results = [];
  results.push(['Sheet Name', 'Last Row', 'Last Column', 'Cell Count', 
                 'Get Values (sec)', 'Get Formulas (sec)', 'Get Formats (sec)', 
                 'Total Load Time (sec)', 'Notes']);
  
  // Measure performance for each sheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    Logger.log("--- Measuring Sheet: " + sheetName + " ---");
    
    try {
      // Get sheet dimensions
      var startTime = new Date().getTime();
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var endTime = new Date().getTime();
      var metadataTime = (endTime - startTime) / 1000;
      
      Logger.log("  Last Row: " + lastRow + ", Last Column: " + lastCol);
      Logger.log("  Metadata retrieval time: " + metadataTime.toFixed(3) + " seconds");
      
      // Skip if sheet is empty
      if (lastRow === 0 || lastCol === 0) {
        Logger.log("  Sheet is empty - skipping performance tests");
        results.push([sheetName, 0, 0, 0, 0, 0, 0, metadataTime.toFixed(3), 'Empty sheet']);
        Logger.log("");
        continue;
      }
      
      // Calculate cell count
      var cellCount = lastRow * lastCol;
      Logger.log("  Estimated cell count: " + cellCount.toLocaleString());
      
      // Measure time to get values
      startTime = new Date().getTime();
      var range = sheet.getRange(1, 1, lastRow, lastCol);
      var values = range.getValues();
      endTime = new Date().getTime();
      var valuesTime = (endTime - startTime) / 1000;
      Logger.log("  Time to get values: " + valuesTime.toFixed(3) + " seconds");
      
      // Measure time to get formulas
      startTime = new Date().getTime();
      var formulas = range.getFormulas();
      endTime = new Date().getTime();
      var formulasTime = (endTime - startTime) / 1000;
      Logger.log("  Time to get formulas: " + formulasTime.toFixed(3) + " seconds");
      
      // Count formulas
      var formulaCount = 0;
      for (var r = 0; r < formulas.length; r++) {
        for (var c = 0; c < formulas[r].length; c++) {
          if (formulas[r][c] !== '') {
            formulaCount++;
          }
        }
      }
      Logger.log("  Formula count: " + formulaCount.toLocaleString());
      
      // Measure time to get formats (this can be slow for large ranges)
      startTime = new Date().getTime();
      var formats = range.getNumberFormats();
      endTime = new Date().getTime();
      var formatsTime = (endTime - startTime) / 1000;
      Logger.log("  Time to get formats: " + formatsTime.toFixed(3) + " seconds");
      
      // Calculate total load time
      var totalTime = valuesTime + formulasTime + formatsTime + metadataTime;
      Logger.log("  TOTAL LOAD TIME: " + totalTime.toFixed(3) + " seconds");
      
      // Store results
      var notes = '';
      if (formulaCount > 0) {
        notes = formulaCount.toLocaleString() + ' formulas';
      }
      if (totalTime > 5) {
        notes += (notes ? ', ' : '') + 'SLOW (>5s)';
      }
      
      results.push([
        sheetName,
        lastRow,
        lastCol,
        cellCount,
        valuesTime.toFixed(3),
        formulasTime.toFixed(3),
        formatsTime.toFixed(3),
        totalTime.toFixed(3),
        notes
      ]);
      
    } catch (e) {
      Logger.log("  ERROR processing sheet: " + e.toString());
      results.push([sheetName, 'ERROR', 'ERROR', 'ERROR', 'ERROR', 'ERROR', 'ERROR', 'ERROR', e.toString()]);
    }
    
    Logger.log("");
  }
  
  // Summary
  Logger.log("=== Performance Measurement Complete ===");
  Logger.log("Total sheets measured: " + sheets.length);
  
  // Calculate summary statistics
  var totalLoadTime = 0;
  var slowSheets = [];
  for (var j = 1; j < results.length; j++) {
    if (results[j][7] !== 'ERROR' && !isNaN(parseFloat(results[j][7]))) {
      var loadTime = parseFloat(results[j][7]);
      totalLoadTime += loadTime;
      if (loadTime > 5) {
        slowSheets.push({name: results[j][0], time: loadTime});
      }
    }
  }
  
  Logger.log("Total load time (all sheets): " + totalLoadTime.toFixed(3) + " seconds");
  
  if (slowSheets.length > 0) {
    Logger.log("Slow sheets (>5 seconds):");
    slowSheets.sort(function(a, b) { return b.time - a.time; });
    for (var k = 0; k < slowSheets.length; k++) {
      Logger.log("  - " + slowSheets[k].name + ": " + slowSheets[k].time.toFixed(3) + " seconds");
    }
  }
  
  // Optionally write results to a sheet
  writeResultsToSheet(ss, results);
}

/**
 * Writes performance results to a "Performance Report" sheet
 * Creates the sheet if it doesn't exist, or clears and updates if it does
 */
function writeResultsToSheet(ss, results) {
  try {
    var reportSheet = ss.getSheetByName("Performance Report");
    
    // Create sheet if it doesn't exist
    if (!reportSheet) {
      reportSheet = ss.insertSheet("Performance Report");
      Logger.log("Created 'Performance Report' sheet");
    } else {
      // Clear existing data
      reportSheet.clear();
      Logger.log("Cleared existing 'Performance Report' sheet");
    }
    
    // Write results
    if (results.length > 0) {
      reportSheet.getRange(1, 1, results.length, results[0].length).setValues(results);
      
      // Format header row
      var headerRange = reportSheet.getRange(1, 1, 1, results[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
      
      // Auto-resize columns
      for (var i = 1; i <= results[0].length; i++) {
        reportSheet.autoResizeColumn(i);
      }
      
      // Sort by total load time (descending) - skip header row
      if (results.length > 2) {
        var dataRange = reportSheet.getRange(2, 1, results.length - 1, results[0].length);
        reportSheet.sort(8, false); // Sort by column 8 (Total Load Time), descending
      }
      
      Logger.log("Results written to 'Performance Report' sheet");
    }
  } catch (e) {
    Logger.log("Could not write results to sheet: " + e.toString());
    Logger.log("Results are still available in the Logger");
  }
}

/**
 * Quick performance test for a specific sheet
 * Useful for testing individual sheets during optimization
 */
function measureSingleSheet(sheetName) {
  const SPREADSHEET_ID = '1q4YqY_vLHsmnN6PDihGAo2seCc0Tugo_5x9QDTF1cek';
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found");
    return;
  }
  
  Logger.log("=== Measuring Sheet: " + sheetName + " ===");
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    Logger.log("Sheet is empty");
    return;
  }
  
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  
  // Test getting values
  var startTime = new Date().getTime();
  var values = range.getValues();
  var endTime = new Date().getTime();
  Logger.log("Get values: " + ((endTime - startTime) / 1000).toFixed(3) + " seconds");
  
  // Test getting formulas
  startTime = new Date().getTime();
  var formulas = range.getFormulas();
  endTime = new Date().getTime();
  Logger.log("Get formulas: " + ((endTime - startTime) / 1000).toFixed(3) + " seconds");
  
  // Count formulas
  var formulaCount = 0;
  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c] !== '') {
        formulaCount++;
      }
    }
  }
  Logger.log("Total formulas: " + formulaCount.toLocaleString());
}

