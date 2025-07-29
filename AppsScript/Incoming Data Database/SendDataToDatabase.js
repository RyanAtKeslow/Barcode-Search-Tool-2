function SendDataToDatabase() {
  Logger.log('ProcessUnprocessedCamerasRows (time-driven) TRIGGERED');

  // Spreadsheet and sheet IDs
  const CAMERAS_SHEET_ID = '13PMB5l5PJr4HHQ0W9A7KvTu2derCVtLKRHtoJ-2LxW4';
  const CAMERA_DATABASE_SHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
  const CAMERA_DATABASE_SHEETS = [
    'ALEXA 35 Body Status',
    'VENICE 2 Body Status',
    'Alexa Mini LF Body Status',
    'Alexa Mini Body Status',
    'Venice 1 Body Status'
  ];

  // Mapping of equipment names to their corresponding sheets
  const CAMERA_NAME_TO_SHEET = {
    'ARRI ALEXA 35 Camera Body': 'ALEXA 35 Body Status',
    'Sony VENICE 2 Digital Camera': 'VENICE 2 Body Status',
    'ARRI ALEXA Mini LF Camera Body': 'Alexa Mini LF Body Status',
    'ARRI ALEXA Mini Camera Body': 'Alexa Mini Body Status',
    'Sony VENICE 1 HFR Digital Camera': 'Venice 1 Body Status'
  };

  // Get the Cameras sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Cameras');
  if (!sheet) {
    Logger.log('Cameras sheet not found. Exiting.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let unprocessedRows = [];
  for (let row = 1; row < data.length; row++) { // Start at 1 to skip header
    const timestamp = data[row][0]; // Column A (index 0)
    const barcode = data[row][1]; // Column B (index 1)
    const equipmentName = data[row][2]; // Column C (index 2)
    const status = data[row][3]; // Column D (index 3)
    const processed = data[row][25]; // Column Z (index 25)

    // Skip empty rows or already processed
    if ((!barcode && !equipmentName && !status) || processed === 'processed') {
      continue;
    }

    unprocessedRows.push({
      rowNum: row + 1, // 1-based for sheet
      timestamp: timestamp,
      barcode: barcode,
      equipmentName: equipmentName,
      status: status
    });
  }

  // Sort unprocessed rows by timestamp (column A)
  unprocessedRows.sort((a, b) => {
    // If either timestamp is missing, treat as later
    if (!a.timestamp) return 1;
    if (!b.timestamp) return -1;
    return new Date(a.timestamp) - new Date(b.timestamp);
  });

  for (const rowObj of unprocessedRows) {
    Logger.log(`Row: ${rowObj.rowNum}, Timestamp: ${rowObj.timestamp}, Status: ${rowObj.status}, Barcode: ${rowObj.barcode}, Equipment: ${rowObj.equipmentName}`);
    if (!rowObj.barcode || !rowObj.equipmentName || !rowObj.status) {
      Logger.log('Missing barcode, equipment name, or status. Skipping row.');
      continue;
    }

    // Fetch location from column G (index 6) of Cameras sheet
    const location = data[rowObj.rowNum - 1][6];
    Logger.log(`Fetched location for barcode ${rowObj.barcode}: ${location}`);

    // Helper to update status and location in external sheet
    function updateStatusAndLocationInExternalSheet(sheetId, sheetName, barcode, status, location) {
      Logger.log(`Attempting to update status and location in external sheet: ${sheetName} (${sheetId}) for barcode: ${barcode}`);
      const extSS = SpreadsheetApp.openById(sheetId);
      const extSheet = extSS.getSheetByName(sheetName);
      if (!extSheet) {
        Logger.log(`Sheet ${sheetName} not found in spreadsheet ${sheetId}`);
        return;
      }
      // For both ARRI and SONY, barcode is column D (4), status is E (5), location is G (7), owner is H (8), etc.
      // We'll update status (E), location (G), and owner (H) as an example. Adjust as needed for your workflow.
      const extData = extSheet.getRange(2, 4, extSheet.getLastRow() - 1, 13).getValues(); // D (barcode) to O (visual) for Alexa 35, D to P for Sony Venice 2
      for (let i = 0; i < extData.length; i++) {
        if (extData[i][0] && extData[i][0].toString().trim() === barcode.toString().trim()) {
          extSheet.getRange(i + 2, 5).setValue(status); // Column E (status)
          extSheet.getRange(i + 2, 7).setValue(location); // Column G (location)
          // Optionally update owner, mount, etc. here if needed
          Logger.log(`Updated status for barcode ${barcode} in ${sheetName} to ${status} and location to ${location}`);
          return;
        }
      }
      Logger.log(`Barcode ${barcode} not found in ${sheetName}`);
    }

    // Find matching sheet based on equipment name
    const matchingSheet = CAMERA_NAME_TO_SHEET[rowObj.equipmentName.trim()];

    if (matchingSheet) {
      Logger.log(`Found matching sheet: ${matchingSheet}`);
      updateStatusAndLocationInExternalSheet(CAMERA_DATABASE_SHEET_ID, matchingSheet, rowObj.barcode, rowObj.status, location);
    } else {
      Logger.log(`No matching camera database sheet found for equipment name: "${rowObj.equipmentName}". No action taken.`);
    }

    // Mark the row as processed in column Z and set background to light green
    sheet.getRange(rowObj.rowNum, 26).setValue('processed');
    sheet.getRange(rowObj.rowNum, 1, 1, sheet.getLastColumn()).setBackground('#b7e1cd');
    Logger.log(`Row ${rowObj.rowNum} marked as processed in column Z and background set to light green.`);
  }
} 