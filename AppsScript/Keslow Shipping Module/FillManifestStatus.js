function fillManifestStatus() {
  Logger.log('Starting fillManifestStatus script');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var manifestSheet = ss.getActiveSheet();
  var dbSheet = ss.getSheetByName('Database');
  if (!dbSheet) {
    Logger.log('❌ Database sheet not found.');
    SpreadsheetApp.getUi().alert('Database sheet not found.');
    return;
  }

  // Get all barcodes from F12 down (was E12)
  var startRow = 12;
  // Get both column A (checkboxes) and column F (barcodes)
  var manifestData = manifestSheet.getRange(startRow, 1, manifestSheet.getLastRow() - startRow + 1, 6).getValues();
  Logger.log('Fetched ' + manifestData.length + ' rows from manifest sheet (columns A and F)');

  // Filter to only process rows where column A is FALSE and column F has a non-empty barcode
  var barcodes = manifestData
    .filter(row => row[0] === false && row[5] !== '') // Column A is FALSE and column F is not empty
    .map(row => row[5]); // Get barcode from column F
  Logger.log('Found ' + barcodes.length + ' non-empty barcodes to process (where column A is FALSE)');

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

  // Process each barcode
  manifestData.forEach((row, i) => {
    var barcode = String(row[5]); // Convert barcode to string
    if (!barcode || row[0] !== false) { // Skip if no barcode or column A is not FALSE
      return;
    }

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

  // Write results to column J (col 10), starting at J12
  manifestSheet.getRange(startRow, 10, results.length, 1).setValues(results);
  Logger.log('Wrote ' + processedCount + ' results to column J, starting at J' + startRow);

  // --- Create Gmail draft with .csv attachment ---
  try {
    Logger.log('Preparing to create Gmail draft with manifest CSV...');
    // Get all data from B5 downwards (until the last row with data in any column)
    var firstDataRow = 5;
    var lastRow = manifestSheet.getLastRow();
    var lastCol = manifestSheet.getLastColumn();
    var data = manifestSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, lastCol - 1).getValues();
    Logger.log('Fetched data for CSV: rows ' + firstDataRow + ' to ' + lastRow + ', columns 2 to ' + lastCol);

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

    // Fetch Origin, Destination, Contract #
    var origin = manifestSheet.getRange('B8').getValue();
    var destination = manifestSheet.getRange('D8').getValue();
    var contractNum = manifestSheet.getRange('C14').getValue();

    // Translate cities to three-letter codes
    var translatedOrigin = translateCity(origin);
    var translatedDestination = translateCity(destination);
    Logger.log('Translated cities - Origin: ' + origin + ' -> ' + translatedOrigin + ', Destination: ' + destination + ' -> ' + translatedDestination);

    // Compose email subject and body
    var emailTitle = 'GT ' + contractNum + ' (' + translatedOrigin + ')->(' + translatedDestination + ')';
    var emailBody = "I've attached a CSV with the manifest for " + emailTitle + " below. See column J for the previous shipping history.";
    var csvFileName = emailTitle + ".csv";
    Logger.log('CSV file name: ' + csvFileName);

    // Create blob for attachment
    var blob = Utilities.newBlob(csv, 'text/csv').setName(csvFileName);

    // Get user email
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log('User email for draft: ' + userEmail);

    // Create Gmail draft
    GmailApp.createDraft(userEmail, emailTitle, emailBody, {attachments: [blob]});
    Logger.log('Gmail draft created for ' + userEmail + ' with ' + csvFileName + ' attached.');
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

    // Get data from B5:AA of manifest sheet
    var lastRow = manifestSheet.getLastRow();
    var data = manifestSheet.getRange(5, 2, lastRow - 4, 26).getValues();
    Logger.log('Fetched data from B5:AA of manifest sheet');

    // Write data to new sheet starting at A5
    newSheet.getRange(5, 1, data.length, data[0].length).setValues(data);
    Logger.log('Wrote data to new sheet starting at A5');

    // Format the new sheet to use overflow instead of wrap
    newSheet.getRange('A5:Z' + (data.length + 4)).setWrap(false);
    Logger.log('Formatted new sheet with overflow');

    // Clear B5:O in the manifest sheet
    manifestSheet.getRange('B5:O').clearContent();
    Logger.log('Cleared B5:O in manifest sheet');
  } catch (e) {
    Logger.log('❌ Error creating new sheet: ' + e.toString());
  }

  SpreadsheetApp.getUi().alert('Manifest status filled in column J, missing barcodes added to database, and email draft created.');
  Logger.log('fillManifestStatus script complete');
} 