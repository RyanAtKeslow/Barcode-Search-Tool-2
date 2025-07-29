/**
 * Shipping Database Lookup
 * Searches through specified sheets for barcode information and shipment history
 */

function searchForBarcodesAndHistory() {
  console.log('Starting barcode search and history update process...');
  
  const EQUIPMENT_SCHEDULE_ID = '1jMMJduxpx4cpGG1t7LdsdGWFwdFlBkD9VCD4GWlLjDo';
  console.log('Using Equipment Schedule ID:', EQUIPMENT_SCHEDULE_ID);
  
  const sheetsToSearch = [
    'Camera',
    'ZEISS',
    'Leitz',
    'Cooke',
    'Ultimate Zoom Tab - NEW',
    'E-EF-FE-B4 Prime/Zooms',
    'Ancient Optics',
    ' Large Format A-O',
    'Large Format S-Z  /  FF Anamorphic',
    'OTHER (Vantage/MiniHawk/SuperBaltar/Kowa/&more)',
    'SPECIALTY',
    'Laowa Lenses',
    'Anamorphic (Super 35)',
    '16mm Format',
    'Director Viewfinders LF',
    'Wireless Follow Focus',
    'Flight Pack',
    'Consignor Use Only',
    'Leitz Loaner - Check w/Zack'
  ];
  console.log('Will search through', sheetsToSearch.length, 'sheets');

  // Get the equipment schedule spreadsheet
  console.log('Opening equipment schedule spreadsheet...');
  const equipmentSpreadsheet = SpreadsheetApp.openById(EQUIPMENT_SCHEDULE_ID);
  console.log('Successfully opened equipment schedule spreadsheet');
  
  console.log('Getting active spreadsheet...');
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  console.log('Active spreadsheet:', activeSpreadsheet.getName());
  
  const databaseSheet = activeSpreadsheet.getSheetByName('Database');
  
  if (!databaseSheet) {
    console.error('Database sheet not found in active spreadsheet');
    throw new Error('Database sheet not found in active spreadsheet');
  }
  console.log('Found Database sheet');

  // Get all barcodes from Database sheet (column C) and filter by 'Keslow Camera' in column F
  console.log('Retrieving barcodes from Database sheet column C and filtering by Keslow Camera in column F...');
  const databaseBarcodes = databaseSheet.getRange('C:C').getValues();
  const databaseOwners = databaseSheet.getRange('F:F').getValues();
  const barcodes = databaseBarcodes.map((row, idx) => {
    const barcode = row[0];
    const owner = databaseOwners[idx][0];
    if (barcode && barcode.toString().trim() !== '' && owner && owner.toString().trim() === 'Keslow Camera') {
      return barcode;
    }
    return null;
  }).filter(bc => bc);
  console.log('Found', barcodes.length, 'barcodes in Database sheet with owner Keslow Camera');

  // Create a map to store barcode to shipment history
  const barcodeToHistory = new Map();

  // Search through each specified sheet
  console.log('Starting search through equipment sheets...');
  sheetsToSearch.forEach(sheetName => {
    console.log('Processing sheet:', sheetName);
    const sheet = equipmentSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      console.warn(`Sheet "${sheetName}" not found in equipment schedule`);
      return;
    }

    // Get all values from column E starting at row 8
    console.log('Retrieving column E data from', sheetName, 'starting at row 8');
    const lastRow = sheet.getLastRow();
    const range = lastRow >= 8 ? sheet.getRange(8, 5, lastRow - 7, 1) : null;
    if (!range) {
      console.log('No data found in column E for sheet', sheetName);
      return;
    }
    const columnE = range.getValues();
    console.log('Found', columnE.length, 'rows in column E');
    
    // Search for barcodes and shipment history
    let barcodesFoundInSheet = 0;
    columnE.forEach((cell, index) => {
      const cellValue = cell[0].toString();
      // Look for BC# pattern and extract only the digits after BC#
      const bcMatch = cellValue.match(/BC# ?(\d+)/g);
      if (bcMatch) {
        bcMatch.forEach(bcString => {
          const barcodeMatch = bcString.match(/BC# ?(\d+)/);
          if (!barcodeMatch) return;
          const barcode = barcodeMatch[1];

          // Only process if barcode is an exact match to one in the filtered database list
          if (!barcodes.includes(barcode)) {
            return;
          }

          // Exclude if NBCA, NBCA*, or NBCA** is present
          if (/NBCA(\*\*|\*|)/.test(cellValue)) {
            return;
          }

          // Extract all parenthetical content after the barcode (e.g., (NBCA), (NBCA**))
          const parenMatches = [...cellValue.matchAll(/\(([^)]+)\)/g)].map(m => m[1]);
          // Extract all '1st sent to ...' phrases
          const sentToMatches = [...cellValue.matchAll(/1st sent to [^,)]+/g)].map(m => m[0]);

          let history = [];
          if (parenMatches.length > 0) history.push(...parenMatches);
          if (sentToMatches.length > 0) history.push(...sentToMatches);

          // Check for date in DD/MM/YY or DD/MM/YYYY format and filter by year
          let hasDate = false;
          let allDatesValid = true;
          const dateRegex = /\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/g;
          let match;
          while ((match = dateRegex.exec(cellValue)) !== null) {
            hasDate = true;
            let year = match[3];
            if (year.length === 2) {
              year = parseInt(year, 10) < 50 ? '20' + year : '19' + year; // crude Y2K logic
            }
            year = parseInt(year, 10);
            if (year < 2024) {
              allDatesValid = false;
              break;
            }
          }
          // If there is no date, skip
          if (!hasDate) {
            return;
          }
          // If there is a date and any year is before 2024, skip
          if (hasDate && !allDatesValid) {
            return;
          }

          if (history.length > 0) {
            console.log('Found history for barcode', barcode, ':', history.join(', '));
            barcodeToHistory.set(barcode, history.join(', '));
          }
        });
      }
    });
    console.log('Found', barcodesFoundInSheet, 'barcodes in sheet', sheetName);
  });

  // Update the Database sheet with found history
  console.log('Updating Database sheet with found history...');
  let updatesMade = 0;
  barcodes.forEach((barcode, index) => {
    if (barcodeToHistory.has(barcode)) {
      databaseSheet.getRange(index + 1, 10).setValue(barcodeToHistory.get(barcode));
      updatesMade++;
    }
  });
  console.log('Updated', updatesMade, 'records in Database sheet');
  console.log('Process completed successfully');
}

// Add menu item to run the function
function onOpen() {
  console.log('Setting up Shipping Tools menu...');
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipping Tools')
    .addItem('Search Barcodes and Update History', 'searchForBarcodesAndHistory')
    .addToUi();
  console.log('Shipping Tools menu created successfully');
} 