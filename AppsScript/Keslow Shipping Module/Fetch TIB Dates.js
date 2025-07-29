function fetchTIBDates() {
  // Logger: Start
  Logger.log('Starting Fetch TIB Dates script');

  // Spreadsheet and sheet IDs
  const DASHBOARD_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const DASHBOARD_SHEET = 'Shipping Dashboard';
  const ESC_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';

  // Column mapping for Shipping Dashboard (row 11 headers)
  const COLS = {
    TIB_YN: 1, // A
    ASSET_ID: 2, // B
    CONTRACT: 3, // C
    QTY: 4, // D
    SERIAL: 5, // E
    BARCODE: 6, // F
    TRANSACTION: 7, // G
    CATEGORY: 8, // H
    DESCRIPTION: 9, // I
    HARMONIZED: 10, // J
    ORIGIN: 11, // K
    CASE: 12, // L
    VALUE: 13, // M
    DIMENSIONS: 14, // N
    WEIGHT: 15, // O
    TIB_DATE: 16, // P (write result)
    TIB_MONTHS: 17 // Q (write result)
  };

  // Get dashboard sheet and barcodes
  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET);
  const dashboardData = dashboardSheet.getDataRange().getValues();
  Logger.log('Loaded Shipping Dashboard data');

  // Get all barcodes from column F (BARCODE), starting from row 12, only if TIB Y/N (column A) is TRUE
  const barcodes = dashboardData.slice(11)
    .filter(row => row[COLS.TIB_YN - 1] === true)
    .map(row => row[COLS.BARCODE - 1])
    .filter(bc => bc && bc !== '');
  Logger.log('Barcodes to process: ' + barcodes.join(', '));

  // Open ESC spreadsheet and get all sheet names
  const escSpreadsheet = SpreadsheetApp.openById(ESC_ID);
  const escSheets = escSpreadsheet.getSheets();
  Logger.log('Loaded ESC spreadsheet and found ' + escSheets.length + ' sheets');

  // Helper function to normalize date to next 30-day increment
  function normalizeReturnDate(returnDate, today) {
    const returnDateObj = new Date(returnDate);
    const todayObj = new Date(today);
    
    // Get the day of month from today
    const targetDay = todayObj.getDate();
    
    // If return date is before today, start from today
    if (returnDateObj < todayObj) {
      returnDateObj.setTime(todayObj.getTime());
    }
    
    // If return date's day is after target day, move to next month
    if (returnDateObj.getDate() > targetDay) {
      returnDateObj.setMonth(returnDateObj.getMonth() + 1);
    }
    
    // Set to target day
    returnDateObj.setDate(targetDay);
    
    return returnDateObj;
  }

  // Today's date in M/D/YYYY format
  const today = new Date();
  const formattedToday = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

  // Collect results for TIB Dashboard appending
  const tibDashboardRows = [];

  // For each barcode, process
  barcodes.forEach((barcode, idx) => {
    Logger.log('Processing barcode: ' + barcode);
    let found = false;
    let foundSheet, foundRow, foundCol, foundContract, foundContractCol, foundContractDate, foundReturnDate, foundReturnMonths;
    const contractNumber = dashboardData[idx + 11][COLS.CONTRACT - 1];
    let dashboardRow = dashboardData[idx + 11];
    let description = dashboardRow[COLS.DESCRIPTION - 1];
    let serial = dashboardRow[COLS.SERIAL - 1];
    let assetId = dashboardRow[COLS.ASSET_ID - 1];
    let category = dashboardRow[COLS.CATEGORY - 1];
    let contract = dashboardRow[COLS.CONTRACT - 1];
    let tibDate = '';
    let tibMonths = '';
    let tibDays = '';

    // Search each sheet in ESC
    for (const sheet of escSheets) {
      const escData = sheet.getDataRange().getValues();
      // Find barcode in column E (index 4)
      for (let r = 0; r < escData.length; r++) {
        if (escData[r][4] && escData[r][4].toString().includes(barcode)) {
          found = true;
          foundSheet = sheet.getName();
          foundRow = r;
          Logger.log(`Found barcode ${barcode} in ESC sheet '${foundSheet}' at row ${foundRow + 1}`);
          // Scan right from barcode cell to find contract #
          let contractCol = null, contractDate = null;
          for (let c = 5; c < escData[foundRow].length; c++) { // start after barcode col (E, index 4)
            if (escData[foundRow][c] && escData[foundRow][c].toString().includes(contractNumber.toString())) {
              contractCol = c;
              contractDate = escData[0][c];
              break;
            }
          }
          if (contractCol === null) {
            Logger.log('No contract # found to the right for barcode ' + barcode);
            break;
          }
          Logger.log(`Found contract # ${contractNumber} at column ${contractCol + 1}, date ${contractDate}`);
          // Find next non-empty cell to the right
          let nextVal = null, nextCol = null;
          for (let c = contractCol + 1; c < escData[foundRow].length; c++) {
            if (escData[foundRow][c] && escData[foundRow][c] !== '') {
              nextVal = escData[foundRow][c];
              nextCol = c;
              break;
            }
          }
          if (!nextVal) {
            Logger.log('No next non-empty cell found for barcode ' + barcode);
            break;
          }
          Logger.log(`Next non-empty cell found at column ${nextCol + 1}, value: ${nextVal}`);
          // Find next white background cell to the right for return date
          let returnDate = null;
          const escSheet = sheet;
          // Get backgrounds for the barcode's row (not header row)
          const rowBackgrounds = escSheet.getRange(foundRow + 1, 1, 1, escSheet.getLastColumn()).getBackgrounds()[0];
          let firstWhiteCol = null;
          for (let c = nextCol + 1; c < rowBackgrounds.length; c++) {
            if (rowBackgrounds[c] === '#ffffff') {
              firstWhiteCol = c;
              Logger.log(`First white cell found at column ${c + 1}, date: ${escData[0][c]}`);
              break;
            }
          }
          if (firstWhiteCol === null) {
            Logger.log('No white cell found for return date for barcode ' + barcode);
            break;
          }

          // City code to country mapping
          const cityCountryMap = {
            'LA': 'US', 'NO': 'US', 'AT': 'US', 'CH': 'US', 'ABQ': 'US',
            'VN': 'CA', 'TO': 'CA'
          };

          // Get current location from E8 and translate to code
          const currentLocationRaw = dashboardSheet.getRange('E8').getValue().toString().trim();
          const currentLocation = translateCityCode(currentLocationRaw);
          const currentCountry = cityCountryMap[currentLocation] || 'UNK';
          Logger.log(`Current location from E8 (translated): ${currentLocation}, country: ${currentCountry}`);

          // Get origin city from B8 in the dashboard and translate to code
          const originCityRaw = dashboardSheet.getRange('B8').getValue().toString().trim();
          const originCity = translateCityCode(originCityRaw);
          Logger.log(`Origin city from dashboard (translated): ${originCity}`);

          // Check for non-empty cells to the right of the first white cell
          let foundInternationalReturn = false;
          let lastNonEmptyCol = null;
          for (let c = firstWhiteCol + 1; c < escData[foundRow].length; c++) {
            const cellVal = escData[foundRow][c];
            if (cellVal && cellVal !== '') {
              lastNonEmptyCol = c;
              Logger.log(`Non-empty cell to right of white cell: '${cellVal}' at column ${c + 1}`);
              const cellStr = cellVal.toString();
              let destCity = null;
              let destCountry = null;
              if (cellStr.includes('GT')) {
                // Look for pattern CITY1 -> CITY2
                const match = cellStr.match(/([A-Z]{2}) -> ([A-Z]{2})/);
                if (match) {
                  const city1 = match[1];
                  const city2 = match[2];
                  Logger.log(`GT found with route: ${city1} -> ${city2}`);
                  destCity = city2;
                  destCountry = cityCountryMap[destCity] || 'UNK';
                } else {
                  Logger.log('GT found but no city route pattern, skipping.');
                  continue;
                }
              } else {
                // Not GT, check first two/three letters for city code
                const cityCode = translateCityCode(cellStr.substring(0, 6));
                Logger.log(`Non-GT cell, city code: ${cityCode}`);
                destCity = cityCode;
                destCountry = cityCountryMap[destCity] || 'UNK';
              }
              // International border crossing check
              if (destCountry !== currentCountry) {
                Logger.log(`International border crossing detected: ${currentCountry} -> ${destCountry}`);
                returnDate = escData[0][c];
                Logger.log(`Return date set to ${returnDate} due to international crossing.`);
                foundInternationalReturn = true;
                break;
              } else {
                Logger.log(`No international crossing: ${currentCountry} -> ${destCountry}, continue searching.`);
              }
            }
          }
          // If no international return found, fallback to next white cell after last non-empty cell
          if (!foundInternationalReturn && lastNonEmptyCol !== null) {
            let fallbackReturnCol = null;
            for (let c = lastNonEmptyCol + 1; c < rowBackgrounds.length; c++) {
              if (rowBackgrounds[c] === '#ffffff') {
                fallbackReturnCol = c;
                break;
              }
            }
            if (fallbackReturnCol !== null) {
              returnDate = escData[0][fallbackReturnCol];
              Logger.log(`No international return found, fallback to next white cell after last non-empty cell: ${returnDate}`);
            } else {
              returnDate = escData[0][firstWhiteCol];
              Logger.log(`No further white cell found, fallback to first white cell date: ${returnDate}`);
            }
          } else if (!foundInternationalReturn) {
            returnDate = escData[0][firstWhiteCol];
            Logger.log(`No non-empty cells found after white cell, fallback to first white cell date: ${returnDate}`);
          }

          // Calculate months between contractDate and returnDate
          const start = new Date(contractDate);
          const end = new Date(returnDate);
          let months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
          if (end.getDate() > start.getDate()) months++;
          Logger.log(`Months in Canada: ${months}`);

          // Calculate days between today and return date
          const returnDateObj = new Date(returnDate);
          const normalizedReturnDate = normalizeReturnDate(returnDateObj, today);
          const daysDiff = Math.ceil((normalizedReturnDate - today) / (1000 * 60 * 60 * 24));
          Logger.log(`Original return date: ${returnDate}, Normalized return date: ${normalizedReturnDate.toLocaleDateString()}, Days until return: ${daysDiff}`);

          // Write to dashboard
          dashboardSheet.getRange(idx + 12, COLS.TIB_DATE, 1, 2).setValues([[normalizedReturnDate.toLocaleDateString(), months]]);
          Logger.log(`Wrote normalized return date and months to dashboard for barcode ${barcode}`);

          // After determining returnDate and months, set tibDate, tibMonths, and tibDays
          tibDate = normalizedReturnDate.toLocaleDateString();
          tibMonths = months;
          tibDays = daysDiff;
          tibDashboardRows.push([
            description,
            barcode,
            serial,
            assetId,
            category,
            contract,
            tibDate,
            tibMonths,
            tibDays
          ]);
          break;
        }
      }
      if (found) break;
    }
    if (!found) {
      Logger.log('Barcode not found in ESC: ' + barcode);
      dashboardSheet.getRange(idx + 12, COLS.TIB_DATE, 1, 2).setValues([['Not found', '']]);
    }
  });

  // Append to TIB Dashboard sheet
  const tibSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TIB Dashboard');
  if (tibSheet && tibDashboardRows.length > 0) {
    tibSheet.getRange(tibSheet.getLastRow() + 1, 1, tibDashboardRows.length, 9).setValues(tibDashboardRows);
    Logger.log(`Appended ${tibDashboardRows.length} rows to TIB Dashboard sheet.`);
  } else if (!tibSheet) {
    Logger.log('TIB Dashboard sheet not found.');
  }

  Logger.log('Fetch TIB Dates script complete');
} 