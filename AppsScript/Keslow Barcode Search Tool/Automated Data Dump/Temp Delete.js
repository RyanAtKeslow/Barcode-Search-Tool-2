function tempDeleteDebug() {
  try {
    Logger.log('🔍 Getting active spreadsheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('❌ Could not access active spreadsheet');
      return;
    }
    Logger.log('✅ Active spreadsheet accessed.');
    Logger.log('🔍 About to get sheet by name...');
    const sheet = ss.getSheetByName('Barcode Dictionary');
    if (sheet) {
      Logger.log('✅ Found sheet: Barcode Dictionary');
    } else {
      Logger.log('❌ Sheet "Barcode Dictionary" not found.');
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
}

function tempDeleteDebugById() {
  try {
    Logger.log('🔍 Getting active spreadsheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('❌ Could not access active spreadsheet');
      return;
    }
    Logger.log('✅ Active spreadsheet accessed.');

    // Replace this with your actual sheet ID (tab ID, a number)
    const targetSheetId = 1966003343; // <-- update this with your actual sheet ID

    Logger.log('🔍 Searching for sheet by ID: ' + targetSheetId);
    const sheets = ss.getSheets();
    let found = false;
    for (const sheet of sheets) {
      if (sheet.getSheetId() === targetSheetId) {
        Logger.log('✅ Found sheet with ID: ' + targetSheetId + ' (Name: ' + sheet.getName() + ')');
        found = true;
        break;
      }
    }
    if (!found) {
      Logger.log('❌ No sheet found with ID: ' + targetSheetId);
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
}

function tempListAllSheets() {
  try {
    Logger.log('🔍 Getting active spreadsheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('❌ Could not access active spreadsheet');
      return;
    }
    Logger.log('✅ Active spreadsheet accessed.');
    const sheets = ss.getSheets();
    Logger.log('Sheet count: ' + sheets.length);
    sheets.forEach(function(sheet) {
      Logger.log('Sheet name: ' + sheet.getName() + ', ID: ' + sheet.getSheetId());
    });
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
  }
} 