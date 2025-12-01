/**
 * Serial Database Sync
 * 
 * This script copies data from the Master Barcode Database Asset Level spreadsheet
 * to the Barcode & Serial Database in Camera Service Forms spreadsheet.
 * 
 * Features:
 * - Copies data in chunks of 9,000 rows to avoid Google Sheets limitations
 * - Runs automatically at 6am PST, Monday through Saturday
 * - Manual refresh via custom menu "Serial Database" > "Refresh Serial Database"
 * - Writes timestamp message in A1 after successful copy
 * 
 * Spreadsheet IDs:
 * - Master: 1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM
 * - Destination: 1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw
 * - Destination Sheet GID: 157146526
 */

// Spreadsheet IDs
const MASTER_SPREADSHEET_ID = '1-P6_duXcx3CSDecKHn-YNMj3VdIKzfVmLDecgXuNqlM';
const DESTINATION_SPREADSHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
const DESTINATION_SHEET_GID = '157146526';

// Copy configuration
const CHUNK_SIZE = 9000; // Copy in chunks of 9,000 rows
const SOURCE_RANGE = 'A2:I500000'; // Source range (A2:I500,000)
const DESTINATION_START_ROW = 2; // Start writing at row 2 (row 1 is for timestamp)

/**
 * Main function to copy database from master to destination
 */
function copySerialDatabase() {
  Logger.log('Starting Serial Database copy process...');
  
  try {
    // Open source and destination spreadsheets
    const masterSpreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const destinationSpreadsheet = SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);
    
    // Get the first sheet from master (gid=0)
    const masterSheet = masterSpreadsheet.getSheets()[0];
    
    // Get the destination sheet by GID
    let destinationSheet = null;
    const sheets = destinationSpreadsheet.getSheets();
    const targetGid = parseInt(DESTINATION_SHEET_GID, 10);
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === targetGid) {
        destinationSheet = sheets[i];
        break;
      }
    }
    
    if (!destinationSheet) {
      Logger.log('ERROR: Could not find destination sheet with GID ' + DESTINATION_SHEET_GID);
      return;
    }
    
    Logger.log('Master sheet: ' + masterSheet.getName());
    Logger.log('Destination sheet: ' + destinationSheet.getName());
    
    // Get the actual data range from master (find last row with data)
    const masterLastRow = masterSheet.getLastRow();
    const masterLastCol = masterSheet.getLastColumn();
    
    Logger.log('Master sheet last row: ' + masterLastRow);
    Logger.log('Master sheet last column: ' + masterLastCol);
    
    // Determine the actual range to copy (A2:I[lastRow])
    const actualLastRow = Math.min(masterLastRow, 500000); // Cap at 500,000 as specified
    const numRowsToCopy = actualLastRow - 1; // Subtract 1 because we start at row 2
    
    Logger.log('Rows to copy: ' + numRowsToCopy);
    
    // Clear destination data starting from row 2 (preserve row 1 for timestamp)
    const destinationLastRow = destinationSheet.getLastRow();
    if (destinationLastRow > 1) {
      Logger.log('Clearing existing data in destination sheet (rows 2-' + destinationLastRow + ')');
      destinationSheet.getRange(2, 1, destinationLastRow - 1, 9).clearContent();
    }
    
    // Copy data in chunks
    let totalRowsCopied = 0;
    const numChunks = Math.ceil(numRowsToCopy / CHUNK_SIZE);
    
    Logger.log('Copying ' + numRowsToCopy + ' rows in ' + numChunks + ' chunks of ' + CHUNK_SIZE + ' rows each');
    
    for (let chunkIndex = 0; chunkIndex < numChunks; chunkIndex++) {
      const startRow = 2 + (chunkIndex * CHUNK_SIZE); // Start at row 2
      const endRow = Math.min(startRow + CHUNK_SIZE - 1, actualLastRow);
      const chunkSize = endRow - startRow + 1;
      
      Logger.log('Copying chunk ' + (chunkIndex + 1) + '/' + numChunks + ': rows ' + startRow + '-' + endRow);
      
      // Get data from master (columns A through I)
      const sourceRange = masterSheet.getRange(startRow, 1, chunkSize, 9);
      const sourceData = sourceRange.getValues();
      
      // Write to destination (starting at row 2)
      const destStartRow = DESTINATION_START_ROW + (chunkIndex * CHUNK_SIZE);
      const destRange = destinationSheet.getRange(destStartRow, 1, chunkSize, 9);
      destRange.setValues(sourceData);
      
      totalRowsCopied += chunkSize;
      Logger.log('Chunk ' + (chunkIndex + 1) + ' completed: ' + chunkSize + ' rows copied');
      
      // Small delay to avoid rate limiting
      Utilities.sleep(100);
    }
    
    Logger.log('Total rows copied: ' + totalRowsCopied);
    
    // Write timestamp message to A1
    const timestamp = new Date();
    const timestampString = timestamp.toLocaleString('en-US', {
      timeZone: 'America/Los_Angeles',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });
    
    const message = 'Asset Database copied on ' + timestampString;
    destinationSheet.getRange(1, 1).setValue(message);
    Logger.log('Timestamp written to A1: ' + message);
    
    Logger.log('Serial Database copy completed successfully!');
    
  } catch (error) {
    Logger.log('ERROR in copySerialDatabase: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    throw error;
  }
}

/**
 * Creates custom menu "Serial Database" with "Refresh Serial Database" option
 * This runs when the spreadsheet is opened
 */
function onOpen(e) {
  createSerialDatabaseMenu_();
}

/**
 * Creates custom menu when add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Builds the custom menu and attaches it to the active spreadsheet UI
 * Uses '_' suffix to avoid accidental exposure as a menu item
 */
function createSerialDatabaseMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Serial Database')
    .addItem('Refresh Serial Database', 'copySerialDatabase')
    .addToUi();
}

/**
 * Sets up time-driven trigger to run copySerialDatabase at 6am PST, Monday through Saturday
 * This function should be run once manually to set up the trigger
 * 
 * Note: The script timezone is set to "America/Los_Angeles" in appsscript.json,
 * so atHour(6) will correctly target 6am PST/PDT
 */
function setupSerialDatabaseTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'copySerialDatabase') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing trigger');
    }
  });
  
  // Create triggers for Monday through Saturday at 6am PST/PDT
  // Since the script timezone is "America/Los_Angeles", atHour(6) will correctly
  // handle both PST (UTC-8) and PDT (UTC-7) automatically
  
  const weekdays = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
    ScriptApp.WeekDay.SATURDAY
  ];
  
  weekdays.forEach(function(weekday) {
    ScriptApp.newTrigger('copySerialDatabase')
      .timeBased()
      .onWeekDay(weekday)
      .atHour(6) // 6am in script timezone (America/Los_Angeles)
      .create();
  });
  
  Logger.log('Time-driven triggers created successfully for Monday-Saturday at 6am PST/PDT');
  Logger.log('Total triggers created: ' + weekdays.length);
}

