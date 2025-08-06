function FetchHarmonyHistory() {
  Logger.log('Starting FetchHarmonyHistory script');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var manifestSheet = ss.getActiveSheet();
  
  // Validate required fields at the very beginning
  var origin = manifestSheet.getRange('B2').getValue();
  var destination = manifestSheet.getRange('D2').getValue();
  
  if (!origin || !destination) {
    var missingFields = [];
    if (!origin) missingFields.push('Origin (B2)');
    if (!destination) missingFields.push('Destination (D2)');
    
    var errorMessage = 'Please fill in the following required fields: ' + missingFields.join(', ');
    Logger.log('❌ Missing required fields: ' + missingFields.join(', '));
    SpreadsheetApp.getUi().alert(errorMessage);
    return;
  }
  
  var dbSheet = ss.getSheetByName('Database');
  if (!dbSheet) {
    Logger.log('❌ Database sheet not found.');
    SpreadsheetApp.getUi().alert('Database sheet not found.');
    return;
  }

  // Find the header row by looking for the specific header pattern
  var headerRow = findHeaderRow(manifestSheet);
  if (headerRow === -1) {
    Logger.log('❌ Header row not found.');
    SpreadsheetApp.getUi().alert('Header row with expected format not found. Looking for header starting with "Asset_ID" in column B.');
    return;
  }
  
  Logger.log('Found header row at: ' + headerRow);
  var startRow = headerRow + 1; // Data starts one row below header

  // Get all barcodes from column F (Barcode #) starting from the row after header
  // Get both column A (checkboxes) and column F (barcodes)
  var lastRow = manifestSheet.getLastRow();
  if (startRow > lastRow) {
    Logger.log('No data rows found below header.');
    SpreadsheetApp.getUi().alert('No data found below the header row.');
    return;
  }
  
  var manifestData = manifestSheet.getRange(startRow, 1, lastRow - startRow + 1, 15).getValues(); // A to O columns
  Logger.log('Fetched ' + manifestData.length + ' rows from manifest sheet (columns A to O)');

  // Get gear transfer number from column C of the first data row
  var gearTransferNum = '';
  if (manifestData.length > 0 && manifestData[0][2]) {
    gearTransferNum = manifestData[0][2]; // Column C (Contract #)
  }
  
  // Validate that gear transfer number exists
  if (!gearTransferNum) {
    Logger.log('❌ Gear Transfer Number not found in column C of first data row.');
    SpreadsheetApp.getUi().alert('Gear Transfer Number not found in column C of the first data row below the header.');
    return;
  }
  
  Logger.log('Found Gear Transfer Number: ' + gearTransferNum);

  // Filter to only process rows where:
  // 1. Column A is FALSE 
  // 2. Column F has a non-empty barcode
  // 3. At least one cell in B:O range is not empty
  var filteredData = manifestData.filter(row => {
    // Check if column A is FALSE and column F has a barcode
    if (row[0] !== false || !row[5]) return false;
    
    // Check if at least one cell in B:O range (indices 1-14) is not empty
    var hasDataInRange = false;
    for (var i = 1; i <= 14; i++) {
      if (row[i] !== null && row[i] !== undefined && row[i] !== '') {
        hasDataInRange = true;
        break;
      }
    }
    return hasDataInRange;
  });
  
  var barcodes = filteredData.map(row => row[5]); // Get barcode from column F
  Logger.log('Found ' + barcodes.length + ' valid rows to process (column A is FALSE, has barcode, and has data in B:O range)');

  // Convert all barcodes to strings before processing
  barcodes = barcodes.map(String);
  Logger.log('Converted barcodes to strings');

  // Get all barcodes from Database column C
  var dbBarcodes = dbSheet.getRange(2, 3, dbSheet.getLastRow() - 1, 1).getValues().flat();
  Logger.log('Fetched ' + dbBarcodes.length + ' barcodes from Database column C');

  // Get reference to Shipping Dashboard sheet
  var shippingDashboard = ss.getSheetByName('Shipping Dashboard');
  if (!shippingDashboard) {
    Logger.log('❌ Shipping Dashboard sheet not found.');
    SpreadsheetApp.getUi().alert('Shipping Dashboard sheet not found.');
    return;
  }

  // Create array to store results, maintaining original row positions
  var results = new Array(manifestData.length).fill(['']);
  var processedCount = 0;

  // Process each barcode using the filtered data
  manifestData.forEach((row, i) => {
    var barcode = String(row[5]); // Convert barcode to string
    
    // Skip if this row doesn't meet our filtering criteria
    if (row[0] !== false || !barcode) return;
    
    // Check if at least one cell in B:O range is not empty
    var hasDataInRange = false;
    for (var j = 1; j <= 14; j++) {
      if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
        hasDataInRange = true;
        break;
      }
    }
    if (!hasDataInRange) return;

    Logger.log('Processing manifest barcode: ' + barcode + ' (type: ' + typeof barcode + ')');
    
    // Find matching row in database
    var dbRowIndex = dbBarcodes.indexOf(barcode);
    Logger.log('Database lookup result: ' + (dbRowIndex !== -1 ? 'Found at row ' + (dbRowIndex + 2) : 'Not found'));
    
    var result = 'No Match';
    if (dbRowIndex !== -1) {
      // Get H and I values directly from the database
      var hValue = dbSheet.getRange(dbRowIndex + 2, 8).getValue() || "No History Found";
      var iValue = dbSheet.getRange(dbRowIndex + 2, 9).getValue() || "No History Found";
      
      Logger.log('Found values - H: ' + hValue + ', I: ' + iValue);
      
      // Only show both values if neither is "No History Found"
      if (hValue !== "No History Found" && iValue !== "No History Found") {
        result = hValue + ' | ' + iValue;
      } else if (hValue !== "No History Found") {
        result = hValue;
      } else if (iValue !== "No History Found") {
        result = iValue;
      } else {
        result = "No History Found";
      }
      
      Logger.log('Row ' + (startRow + i) + ': Barcode ' + barcode + ' => ' + result);
      processedCount++;
    } else {
      // Barcode not found in database - add it from Shipping Dashboard data
      Logger.log('Row ' + (startRow + i) + ': Barcode ' + barcode + ' not found in Database, attempting to add...');
      
      // Get all data from Shipping Dashboard sheet
      var dashboardData = shippingDashboard.getDataRange().getValues();
      var dashboardBarcodes = dashboardData.map(row => String(row[5])); // Column F (index 5)
      
      // Find barcode in Shipping Dashboard
      var dashboardRowIndex = dashboardBarcodes.indexOf(barcode);
      
      if (dashboardRowIndex !== -1) {
        var category = dashboardData[dashboardRowIndex][7] || ''; // Column H (index 7)
        var equipmentName = dashboardData[dashboardRowIndex][8] || ''; // Column I (index 8)
        
        Logger.log('Found in Shipping Dashboard - Equipment: ' + equipmentName + ', Category: ' + category);
        
        // Add new row to database sheet
        var newRowIndex = dbSheet.getLastRow() + 1;
        dbSheet.getRange(newRowIndex, 2).setValue(equipmentName); // Column B
        dbSheet.getRange(newRowIndex, 3).setValue(barcode); // Column C
        dbSheet.getRange(newRowIndex, 4).setValue(category); // Column D
        
        result = 'Added to Database';
        Logger.log('Row ' + (startRow + i) + ': Added barcode ' + barcode + ' to database at row ' + newRowIndex);
        processedCount++;
      } else {
        result = 'No Match Found';
        Logger.log('Row ' + (startRow + i) + ': Barcode ' + barcode + ' not found in Shipping Dashboard either');
      }
    }
    results[i] = [result];
  });

  // Write results to column J (col 10), starting from the first data row
  manifestSheet.getRange(startRow, 10, results.length, 1).setValues(results);
  Logger.log('Wrote ' + processedCount + ' results to column J, starting at J' + startRow);

  // --- Create Gmail draft with .csv attachment ---
  try {
    Logger.log('Preparing to create Gmail draft with manifest CSV...');
    // Get all data from the header row downwards (until the last row with data in any column)
    var lastRow = manifestSheet.getLastRow();
    var lastCol = manifestSheet.getLastColumn();
    var data = manifestSheet.getRange(headerRow, 2, lastRow - headerRow + 1, lastCol - 1).getValues();
    Logger.log('Fetched data for CSV: rows ' + headerRow + ' to ' + lastRow + ', columns 2 to ' + lastCol);

    // Convert to CSV
    var csv = data.map(function(row) {
      return row.map(function(cell) {
        if (cell === null || cell === undefined) return '';
        var str = cell.toString();
        if (str.match(/[",\n]/)) {
          str = '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',');
    }).join('\n');
    Logger.log('CSV content prepared, length: ' + csv.length);

    // Use the validated values from the beginning of the script

    // Translate cities to two-letter codes
    var translatedOrigin = translateCityDropDowns(origin);
    var translatedDestination = translateCityDropDowns(destination);
    Logger.log('Translated cities - Origin: ' + origin + ' -> ' + translatedOrigin + ', Destination: ' + destination + ' -> ' + translatedDestination);

    // Compose email subject and body
    var emailTitle = 'GT ' + gearTransferNum + ' (' + translatedOrigin + ')->(' + translatedDestination + ')';
    var emailBody = "I've attached a CSV with the manifest for " + emailTitle + " below. See column J for the previous shipping history.";
    var csvFileName = emailTitle + ".csv";
    Logger.log('CSV file name: ' + csvFileName);

    // Create blob for attachment
    var blob = Utilities.newBlob(csv, 'text/csv').setName(csvFileName);

    // Create Gmail draft (no recipient specified - user can add recipients later)
    GmailApp.createDraft('', emailTitle, emailBody, {attachments: [blob]});
    Logger.log('Gmail draft created with ' + csvFileName + ' attached (no recipient specified).');
  } catch (e) {
    Logger.log('❌ Error creating Gmail draft: ' + e.toString());
  }

  // Create new sheet with manifest data
  try {
    Logger.log('Creating new sheet with manifest data...');
    // Create new sheet with emailTitle as name
    var newSheet = ss.insertSheet(emailTitle);
    Logger.log('Created new sheet: ' + emailTitle);

    // Write emailTitle to A4
    newSheet.getRange('A4').setValue(emailTitle);
    Logger.log('Wrote emailTitle to A4');

    // Get data from header row to last row, columns B to O
    var lastRow = manifestSheet.getLastRow();
    var data = manifestSheet.getRange(headerRow, 2, lastRow - headerRow + 1, 14).getValues(); // B to O (14 columns)
    Logger.log('Fetched data from header row of manifest sheet');

    // Write data to new sheet starting at A5
    newSheet.getRange(5, 1, data.length, data[0].length).setValues(data);
    Logger.log('Wrote data to new sheet starting at A5');

    // Format the new sheet to use overflow instead of wrap
    newSheet.getRange('A5:O' + (data.length + 4)).setWrap(false);
    Logger.log('Formatted new sheet with overflow');

    // Clear data area in the manifest sheet (from start row to last row, columns B to O)
    manifestSheet.getRange(startRow, 2, lastRow - startRow + 1, 14).clearContent();
    Logger.log('Cleared data area in manifest sheet');
  } catch (e) {
    Logger.log('❌ Error creating new sheet: ' + e.toString());
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Manifest status filled in column J, missing barcodes added to database, and email draft created.', 'Process Complete', 5);
  Logger.log('fillManifestStatus script complete');
}

/**
 * Helper function to find the header row by looking for the expected header pattern
 * Returns the row number (1-based) or -1 if not found
 */
function findHeaderRow(sheet) {
  var maxRowsToSearch = 50; // Reasonable limit to search for header
  var lastRow = Math.min(sheet.getLastRow(), maxRowsToSearch);
  
  // Expected header values for columns B through O
  var expectedHeaders = [
    "Asset_ID", "Contract #", "Qty", "Serial #", "Barcode #", 
    "Asset Transaction Code", "Equipment Category", "Description", 
    "Equipment Harmonized Code", "Origin Country", "Case #", 
    "Value", "Dimensions", "Weight"
  ];
  
  for (var row = 1; row <= lastRow; row++) {
    // Get values from columns B to O for this row
    var rowData = sheet.getRange(row, 2, 1, 14).getValues()[0];
    
    // Check if this row matches our expected header pattern
    var isHeaderRow = true;
    for (var col = 0; col < expectedHeaders.length; col++) {
      var cellValue = String(rowData[col]).trim();
      if (cellValue !== expectedHeaders[col]) {
        isHeaderRow = false;
        break;
      }
    }
    
    if (isHeaderRow) {
      Logger.log('Found header row at row ' + row);
      return row;
    }
  }
  
  Logger.log('Header row not found in first ' + maxRowsToSearch + ' rows');
  return -1;
} 