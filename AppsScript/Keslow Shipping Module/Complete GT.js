function completeGT() {
  Logger.log('Starting completeGT script');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var manifestSheet = ss.getActiveSheet();
  var sheetName = manifestSheet.getName();
  
  // Show confirmation dialog
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Complete GT Confirmation',
    'You are about to Complete the GT "' + sheetName + '". If this is correct, click "Yes". This will DELETE this sheet when finished, please make sure everything has been accounted for',
    ui.ButtonSet.YES_NO
  );

  // Check user's response
  if (response !== ui.Button.YES) {
    Logger.log('User cancelled the operation');
    return;
  }

  var dbSheet = ss.getSheetByName('Database');
  
  if (!dbSheet) {
    Logger.log('❌ Database sheet not found.');
    SpreadsheetApp.getUi().alert('Database sheet not found.');
    return;
  }

  // Get origin and destination from A8 and C8
  var origin = manifestSheet.getRange('A8').getValue();
  var destination = manifestSheet.getRange('C8').getValue();
  Logger.log('Original Origin: ' + origin + ', Destination: ' + destination);

  origin = translateCity(origin);
  destination = translateCity(destination);
  Logger.log('Translated Origin: ' + origin + ', Destination: ' + destination);
  Logger.log('Will check for CA updates if destination is VAN/TOR and origin is not: ' + ((destination === 'VAN' || destination === 'TOR') && origin !== 'VAN' && origin !== 'TOR'));
  Logger.log('Will check for US updates if destination is not VAN/TOR and origin is VAN/TOR: ' + ((destination !== 'VAN' && destination !== 'TOR') && (origin === 'VAN' || origin === 'TOR')));

  // Get all barcodes from E14 down
  var startRow = 14;
  var barcodes = manifestSheet.getRange(startRow, 5, manifestSheet.getLastRow() - startRow + 1, 1).getValues().flat();
  Logger.log('Fetched ' + barcodes.length + ' barcodes from column E, starting at E' + startRow);
  Logger.log('First few manifest barcodes: ' + barcodes.slice(0, 5).join(', '));

  // Get all barcodes from Database column C
  var dbBarcodes = dbSheet.getRange(1, 3, dbSheet.getLastRow(), 1).getValues().flat();
  Logger.log('Fetched ' + dbBarcodes.length + ' barcodes from Database column C');
  Logger.log('First few database barcodes: ' + dbBarcodes.slice(0, 5).join(', '));

  // Convert all barcodes to strings for consistent comparison
  barcodes = barcodes.map(b => b ? b.toString() : '');
  dbBarcodes = dbBarcodes.map(b => b ? b.toString() : '');
  Logger.log('Converted barcodes to strings for comparison');

  // Get today's date in DD/MM/YY format
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yy");
  Logger.log('Today\'s date: ' + formattedDate);

  // Initialize counters
  var updatedCount = 0;
  var unchangedCount = 0;
  var unchangedHistory = [];

  // Process each barcode
  var results = [];
  for (var i = 0; i < barcodes.length; i++) {
    var barcode = barcodes[i];
    if (!barcode) {
      results.push(['']);
      Logger.log('Row ' + (startRow + i) + ': Empty barcode, skipping');
      continue;
    }

    Logger.log('Processing barcode: ' + barcode + ' (type: ' + typeof barcode + ')');
    
    // Find matching row in database
    var dbRowIndex = dbBarcodes.indexOf(barcode);
    Logger.log('Database lookup result for ' + barcode + ': ' + (dbRowIndex !== -1 ? 'Found at row ' + (dbRowIndex + 1) : 'Not found'));
    
    if (dbRowIndex !== -1) {
      // Get current H and I values
      var currentH = dbSheet.getRange(dbRowIndex + 1, 8).getValue();
      var currentI = dbSheet.getRange(dbRowIndex + 1, 9).getValue();
      Logger.log('Current values - H: ' + currentH + ', I: ' + currentI);
      
      var result = '';
      
      // Check if destination is VAN or TOR and origin is not
      Logger.log('Evaluating conditions for barcode ' + barcode + ': destination=' + destination + ', origin=' + origin);
      if ((destination === 'VAN' || destination === 'TOR') && 
          origin !== 'VAN' && origin !== 'TOR') {
        Logger.log('Condition met: Going to CA (destination is VAN/TOR, origin is not)');
        Logger.log('Current H value: "' + currentH + '" (empty check: ' + (!currentH) + ')');
        if (!currentH) {
          result = "1st sent to CA on " + formattedDate;
          // Update database column H
          Logger.log('Attempting to update database row ' + (dbRowIndex + 1) + ' for barcode: ' + barcode);
          dbSheet.getRange(dbRowIndex + 1, 8).setValue("1st sent to CA on " + formattedDate); // Column H
          Logger.log('Updated database column H for barcode ' + barcode + ' with message: 1st sent to CA on ' + formattedDate);
          updatedCount++;
          Logger.log('Row ' + (startRow + i) + ': First time to CA for barcode ' + barcode);
        } else {
          Logger.log('H column already has value, not updating: ' + currentH);
          unchangedCount++;
          unchangedHistory.push("Barcode " + barcode + ": " + currentH);
        }
      }
      // Check if destination is not VAN or TOR and origin is VAN or TOR
      else if ((destination !== 'VAN' && destination !== 'TOR') && 
               (origin === 'VAN' || origin === 'TOR')) {
        Logger.log('Condition met: Going to US (destination is not VAN/TOR, origin is VAN/TOR)');
        Logger.log('Current I value: "' + currentI + '" (empty check: ' + (!currentI) + ')');
        if (!currentI) {
          result = "1st sent to US on " + formattedDate;
          // Update database column I
          Logger.log('Attempting to update database row ' + (dbRowIndex + 1) + ' for barcode: ' + barcode);
          dbSheet.getRange(dbRowIndex + 1, 9).setValue("1st sent to US on " + formattedDate); // Column I
          Logger.log('Updated database column I for barcode ' + barcode + ' with message: 1st sent to US on ' + formattedDate);
          updatedCount++;
          Logger.log('Row ' + (startRow + i) + ': First time to US for barcode ' + barcode);
        } else {
          Logger.log('I column already has value, not updating: ' + currentI);
          unchangedCount++;
          unchangedHistory.push("Barcode " + barcode + ": " + currentI);
        }
      } else {
        Logger.log('No condition met - not a cross-border shipment or conditions not satisfied');
      }
    } else {
      Logger.log('Row ' + (startRow + i) + ': Barcode ' + barcode + ' not found in Database');
    }
    results.push([result]);
  }

  // Write results to column I (col 9), starting at I12
  manifestSheet.getRange(startRow, 9, results.length, 1).setValues(results);
  Logger.log('Wrote ' + results.length + ' results to column I, starting at I' + startRow);

  // Create summary message
  var summaryMessage = 'GT completion process finished.\n\n';
  summaryMessage += 'Updated ' + updatedCount + ' barcodes with shipping history of "1st sent to ' + 
                   (destination === 'VAN' || destination === 'TOR' ? 'CA' : 'US') + ' on ' + formattedDate + '"\n\n';
  
  if (unchangedCount > 0) {
    summaryMessage += unchangedCount + ' barcodes remain unchanged. Already contained shipping history:\n';
    // Show up to 5 examples of unchanged history
    var examplesToShow = Math.min(5, unchangedHistory.length);
    for (var i = 0; i < examplesToShow; i++) {
      summaryMessage += '- ' + unchangedHistory[i] + '\n';
    }
    if (unchangedHistory.length > 5) {
      summaryMessage += '- ... and ' + (unchangedHistory.length - 5) + ' more\n';
    }
  }

  SpreadsheetApp.getUi().alert(summaryMessage);
  Logger.log('completeGT script complete');

  // Delete the sheet after completion
  try {
    ss.deleteSheet(manifestSheet);
    Logger.log('Successfully deleted sheet: ' + sheetName);
  } catch (e) {
    Logger.log('❌ Error deleting sheet: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error deleting sheet: ' + e.toString());
  }
} 