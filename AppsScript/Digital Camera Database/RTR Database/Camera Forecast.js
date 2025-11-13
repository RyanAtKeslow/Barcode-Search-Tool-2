function getCameraForecast() {
  // Color mapping for logging
  const colorNames = {
    '#ff4444': 'Red',
    '#00ffff': 'Light Blue',
    '#ff7171': 'Salmon',
    '#4a86e8': 'Dark Blue',
    '#bdbdbd': 'Grey',
    '#f9ff71': 'Yellow',
    '#66ff75': 'Green',
    '#ffffff': 'White'
  };

  // Helper function to get color name for logging
  function getColorName(hexColor) {
    return colorNames[hexColor] || hexColor;
  }

  // Function to extract order number and job name from scheduling sheet values
  function extractOrderAndJobNameFromSchedule(value) {
    // Convert to string and handle null/undefined values
    const stringValue = (value !== null && value !== undefined) ? value.toString() : '';
    const orderMatch = stringValue.match(/\b\d{6}\b/); // Match a 6-digit order number
    const nameMatch = stringValue.match(/"([^"]+)"/); // Match text within quotes
    const orderNumber = orderMatch ? orderMatch[0] : '';
    const jobName = nameMatch ? nameMatch[1] : '';
    return { orderNumber, jobName };
  }

  // Function to extract order number and job name from prep bay assignments
  function extractOrderAndJobName(jobString) {
    // Convert to string and handle null/undefined values
    const stringValue = (jobString !== null && jobString !== undefined) ? jobString.toString() : '';
    const orderMatch = stringValue.match(/\b(\d+)\b/);
    const nameMatch = stringValue.match(/"([^"]+)"/);
    const orderNumber = orderMatch ? orderMatch[1] : '';
    const jobName = nameMatch ? nameMatch[1] : '';
    return { orderNumber, jobName };
  }

  // Normalize a string by trimming and converting to lowercase
  function normalizeString(str) {
    return str.trim().toLowerCase().replace(/\s+/g, ' ');
  }

  // Helper function to format a date as MM/DD/YYYY
  function formatDate(date) {
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  }

  // Get today's date in M/D/YYYY format (no leading zeros)
  const today = new Date();
  const formattedDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
  Logger.log('Today\'s date: ' + formattedDate);
  Logger.log('Current year: ' + today.getFullYear());

  // Refresh AI:AM section (Prep Bay Block - today, tomorrow, day after tomorrow) before continuing
  try {
    prepBayBlock();
  } catch (err) {
    Logger.log('Error executing prepBayBlock: ' + err);
  }
  
  // Refresh AJ:AM section (tomorrow-prep) before continuing
  try {
    tomorrowDoubleCheck();
  } catch (err) {
    Logger.log('Error executing tomorrowDoubleCheck: ' + err);
  }
  
  // --------------------- NEW: Tomorrow-prep utilities ---------------------
  // Calculate tomorrow's date (start of day)
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1); // midnight tomorrow

  // Day prefixes used in sheet names
  const dayPrefixes = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];

  // Helper: build expected sheet name for a given date
  function expectedSheetName(dateObj) {
    const prefix = dayPrefixes[dateObj.getDay()];
    return `${prefix} ${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
  }

  // Helper: check if two Date objects represent the same calendar day
  function isSameDay(d1, d2) {
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
  }

  // Helper: does a free-text note contain the target MM/DD date?
  function noteContainsDate(note, dateObj) {
    if (!note) return false;
    const matches = note.match(/\d{1,2}\/\d{1,2}/g);
    if (!matches) return false;
    const targetMonth = dateObj.getMonth() + 1;
    const targetDay   = dateObj.getDate();
    return matches.some(d => {
      const [m, day] = d.split('/').map(Number);
      return m === targetMonth && day === targetDay;
    });
  }

  // Normalize order numbers to digits-only string
  const normalizeOrder = val => String(val || '').replace(/[^0-9]/g, '');

  // Build the job date map from the Prep Bay Assignments
  Logger.log('Starting to build the Prep Bay Job Date Map.');
  const prepBaySpreadsheet = SpreadsheetApp.openById('1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ');
  const prepBaySheets = prepBaySpreadsheet.getSheets();
  const jobDateMap = {};
  const sevenDaysFromNow = new Date();
  sevenDaysFromNow.setDate(today.getDate() + 7);

  // Filter out hidden sheets
  const visibleSheets = prepBaySheets.filter(sheet => !sheet.isSheetHidden());

  // Sort sheets by date and select recent sheets
  const sortedSheets = visibleSheets.map(sheet => {
    const sheetName = sheet.getName();
    const dateMatch = sheetName.match(/\w+ (\d+)\/(\d+)/); // Match format like 'Mon 6/2'
    if (dateMatch) {
      const month = parseInt(dateMatch[1], 10) - 1; // JavaScript months are 0-based
      const day = parseInt(dateMatch[2], 10);
      const sheetDate = new Date(today.getFullYear(), month, day);

      // Adjust year if the sheet date is in the past
      if (sheetDate < today) {
        sheetDate.setFullYear(today.getFullYear() + 1);
      }
      return { sheet, sheetDate };
    }
    return null;
  }).filter(item => item !== null);

  // Sort sheets by date in ASCENDING order so earlier days are processed first
  sortedSheets.sort((a, b) => a.sheetDate - b.sheetDate);

  /* Legacy tomorrow-prep collection logic disabled – handled by tomorrowDoubleCheck() */
  if (false) {
    Logger.log('Legacy tomorrow-prep logic inactive.');
  }

  // --------------------- RESTORED: Build jobDateMap for next 7 days ---------------------
  const recentSheets = sortedSheets.filter(item => item.sheetDate >= today && item.sheetDate <= sevenDaysFromNow);

  Logger.log('Processing prep bay assignments from recent sheets...');
  let prepBayCount = 0;
  for (const { sheet, sheetDate } of recentSheets) {
    const sheetName = sheet.getName();
    const data = sheet.getRange('B:J').getValues(); // Job in B, order in C, notes in J
    for (let i = 0; i < data.length; i++) {
      const jobName     = data[i][0];
      const orderNumber = data[i][1];
      const note        = data[i][8]; // Column J
      let jobDate       = '';

      // Default date from tab name (e.g., "Mon 6/2")
      const sheetDateMatch = sheetName.match(/\w+ (\d+)\/(\d+)/);
      if (sheetDateMatch) {
        const month = parseInt(sheetDateMatch[1], 10) - 1; // 0-based
        const day   = parseInt(sheetDateMatch[2], 10);
        const d     = new Date(today.getFullYear(), month, day);
        if (d < today) d.setFullYear(today.getFullYear() + 1);
        jobDate = formatDate(d);
      }

      // Override with earliest date mentioned in a 'prep' note
      if (note && typeof note === 'string' && note.toLowerCase().includes('prep')) {
        const dateMatches = note.match(/\d{1,2}\/\d{1,2}/g);
        if (dateMatches) {
          const earliest = dateMatches.reduce((earliest, cur) => {
            const [m, d] = cur.split('/').map(Number);
            const dt = new Date(today.getFullYear(), m - 1, d);
            return dt < earliest ? dt : earliest;
          }, new Date(today.getFullYear() + 1, 0, 1));
          jobDate = formatDate(earliest);
        }
      }

      if (jobName && orderNumber) {
        jobDateMap[jobName] = { date: jobDate, orderNumber };
        prepBayCount++;
      }
    }
  }
  Logger.log(`Completed prep bay processing: ${prepBayCount} assignments found`);

  // Get the spreadsheet and Camera sheet
  const spreadsheet = SpreadsheetApp.openById('1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw');
  const cameraSheet = spreadsheet.getSheetByName('Camera');

  // Function to translate camera types based on special rules
  function translateCameraType(cameraType) {
    if (cameraType === 'X = Global Shutter Sensor') {
      return 'RED V-RAPTOR XL [X] 8K VV Digital Camera';
    }
    return cameraType;
  }

  // Function to detect if a job starts with a two-letter city abbreviation 
  function hasOutOfTownPrefix(jobString) {
    if (!jobString || typeof jobString !== 'string') return false;
    
    // Match jobs that start with a two-letter abbreviation followed by a space
    // Examples: "AT xxxxx job name", "NY xxxxx job name", "VN xxxxx job name"
    const outOfTownPattern = /^[A-Z]{2}\s/;
    return outOfTownPattern.test(jobString.toString().toUpperCase());
  }

  // Function to check if a colored segment ends with "returns to LA" or similar
  function checkForReturnToLA(rowValues, rowBackgrounds, startCol, maxCols) {
    if (!rowValues || !rowBackgrounds || startCol < 0) return false;
    
    const startColor = rowBackgrounds[startCol];
    if (!startColor || startColor === '#ffffff') return false;
    
    // Find the end of the colored segment
    let endCol = startCol;
    for (let col = startCol + 1; col < Math.min(rowValues.length, maxCols); col++) {
      if (rowBackgrounds[col] === startColor) {
        endCol = col;
      } else {
        break; // Color changed, segment ended
      }
    }
    
    // Check cells in the colored segment for "returns to LA" or similar phrases
    const returnPatterns = [
      /returns?\s*to\s*la/i,
      /back\s*to\s*la/i,
      /return\s*la/i,
      /la\s*return/i
    ];
    
    for (let col = startCol; col <= endCol; col++) {
      const cellValue = rowValues[col];
      if (cellValue && typeof cellValue === 'string') {
        for (const pattern of returnPatterns) {
          if (pattern.test(cellValue)) {
            Logger.log(`Found return to LA pattern in cell at column ${col + 1}: "${cellValue}"`);
            return true;
          }
        }
      }
    }
    
    return false;
  }

  // Gear Transfer Check function - detects if LA is involved in any GT within the row and date is within next 7 days
  function gtCheck(rowValues) {
    const today = new Date();
    const sevenDaysFromNow = new Date();
    sevenDaysFromNow.setDate(today.getDate() + 7);
    
    // Look through all cells in the row for GT patterns
    for (let i = 0; i < rowValues.length; i++) {
      const cellValue = rowValues[i];
      if (!cellValue || typeof cellValue !== 'string') continue;
      
      const cellStr = cellValue.toString().toUpperCase();
      
      // Check if cell contains "GT" and city direction indicators
      if (cellStr.includes('GT') && (cellStr.includes('>') || cellStr.includes('->'))) {
        // Extract city codes around the direction indicators
        // Match patterns like "LA -> AT", "VN>LA", "LA>NY", etc.
        const gtPattern = /(\w{2,3})\s*-?>\s*(\w{2,3})/;
        const match = cellStr.match(gtPattern);
        
        if (match) {
          const origin = match[1];
          const destination = match[2];
          
          // Only process further if LA is involved
          const isLAInvolved = (origin === 'LA' || destination === 'LA');
          
          if (isLAInvolved) {
            // Parse date from the GT entry
            // Look for date patterns like "8/4", "8/6", "12/25", etc.
            const datePattern = /(\d{1,2})\/(\d{1,2})/;
            const dateMatch = cellValue.match(datePattern);
            
            if (dateMatch) {
              const month = parseInt(dateMatch[1], 10) - 1; // JavaScript months are 0-based
              const day = parseInt(dateMatch[2], 10);
              let gtDate = new Date(today.getFullYear(), month, day);
              
              // If the date appears to be in the past, it's actually next year
              // (since sheet data is always forward-looking)
              if (gtDate < today) {
                gtDate = new Date(today.getFullYear() + 1, month, day);
              }
              
              // Check if GT date is within the next 7 days
              if (gtDate >= today && gtDate <= sevenDaysFromNow) {
                Logger.log(`GT Check: LA involved in transfer (${origin} -> ${destination}) on ${dateMatch[0]} - treating as valid LA camera`);
                return true;
              } else {
                Logger.log(`GT Check: LA transfer (${origin} -> ${destination}) on ${dateMatch[0]} is outside 7-day window - not valid`);
              }
            } else {
              Logger.log(`GT Check: LA transfer (${origin} -> ${destination}) found but no valid date parsed from: "${cellValue}"`);
            }
          }
          // Note: Not logging non-LA transfers to reduce log noise
        }
      }
    }
    return false;
  }

  // Find all rows containing "LOS ANGELES" in column A OR gear transfers involving LA
  const data = cameraSheet.getDataRange().getValues();
  const foundLACameras = [];
  const cameraTypeCount = {}; // Object to store count of each camera type
  
  for (let i = 0; i < data.length; i++) {
    const isLACamera = data[i][0] === 'LOS ANGELES';
    const isGTCamera = gtCheck(data[i]);
    
    if (isLACamera || isGTCamera) {
      foundLACameras.push(i + 1); // Adding 1 because getValues() is 0-based but sheet rows are 1-based
    }
  }

  // For each LA camera, find its type
  Logger.log('Processing LA cameras and finding camera types...');
  const cameraTypes = [];
  for (const row of foundLACameras) {
    // Look for first empty cell above
    let typeRow = row - 1; // Start from the row above the LA entry
    while (typeRow >= 0 && data[typeRow][0] !== '') {
      typeRow--;
    }
    // Get the camera type from column E of the first empty cell's row
    const rawCameraType = data[typeRow][4]; // Column E is index 4
    const cameraType = translateCameraType(rawCameraType); // Apply translation rules
    cameraTypes.push({
      row: row,
      type: cameraType
    });
    // Increment count for this camera type
    cameraTypeCount[cameraType] = (cameraTypeCount[cameraType] || 0) + 1;
  }
  // Log a summary of all found camera types and their quantities
  Logger.log('Camera type summary:');
  Object.entries(cameraTypeCount).forEach(([type, count]) => {
    Logger.log(`  ${type}: ${count}`);
  });

  // Find today's date column
  const headerRow = data[0];
  let todayColumn = 0;
  for (let i = 0; i < headerRow.length; i++) {
    const cell = headerRow[i];
    if (cell instanceof Date) {
      if (
        cell.getFullYear() === today.getFullYear() &&
        cell.getMonth() === today.getMonth() &&
        cell.getDate() === today.getDate()
      ) {
        todayColumn = i + 1; // 1-based index for getRange
        Logger.log(`Found today's date in column ${todayColumn} (cell value: ${cell})`);
        break;
      }
    } else if (typeof cell === 'string') {
      // Also check string version in case some columns are text
      const parts = cell.split('/');
      if (parts.length === 3) {
        const m = parseInt(parts[0], 10);
        const d = parseInt(parts[1], 10);
        const y = parseInt(parts[2], 10);
        if (
          y === today.getFullYear() &&
          m === today.getMonth() + 1 &&
          d === today.getDate()
        ) {
          todayColumn = i + 1;
          Logger.log(`Found today's date in column ${todayColumn} (string cell: ${cell})`);
          break;
        }
      }
    }
  }

  if (todayColumn === 0) {
    Logger.log('Could not find today\'s date in header row. Please check the date format in the spreadsheet.');
    Logger.log('Looking for date: ' + formattedDate + ' or Date object.');
    return;
  }
  Logger.log('Today\'s column: ' + todayColumn);

  // Prepare to collect results for writing to sheet
  const outputRows = [];
  const outputBackgrounds = [];
  const rowColorMap = {}; // key: barcode, value: foundColor

  // List of valid camera types (includes translated types)
  const validCameraTypes = [
    'ARRI Amira',
    'Arri Alexa EV High Speed',
    'ARRI ALEXA LF Camera Body',
    'ARRI ALEXA 35 Camera Body',
    'ARRI ALEXA Mini Camera Body',
    'ARRI ALEXA Mini LF Camera Body',
    'RED DSMC2 5K Digital Camera w/ Gemini™ Sensor',
    'RED Komodo 6K Digital Cinema Camera',
    'RED DSMC2 8K Digital Camera w/ Helium™ Sensor',
    'RED DSMC2 8K VV Digital Camera w/ Monstro™ Sensor',
    'RED Ranger DSMC2 8K VV Digital Camera w/ Monstro™ Sensor',
    'RED V-RAPTOR DSMC3 8K VV Digital Camera ',
    'RED V-RAPTOR DSMC3 [X] 8K VV Digital Camera',
    'RED V-RAPTOR XL [X] 8K VV Digital Camera',
    'SONY Burano',
    'SONY VENICE 2',
    'Sony Venice Rialto 2 Mini Camera Body Adapter',
    'Sony Venice Rialto 2 Camera Body Adapter',
    'SONY VENICE 1',
    'Sony FX3 Digital Camera',
    'Sony FX6 Digital Camera',
    'Sony FX9 Digital Camera',
    'Sony A7S II Mirrorless Digital Camera',
    'Sony PDW-F800 XDCAM ',
    'Phantom Flex4K High-Speed Digital Camera',
    'Arri 235 Camera',
    'Arri 35 IIC Hand Crank Camera',
    'Arri 435 ES Camera - Perf Changes not possible',
    'Arri 435 Extreme 3-Perf Camera - Perf Changes not possible',
    'Indieassist 435 HD IVS ',
    'Arricam LT Camera',
    'Arricam LT HD IVS',
    'Indieassist LT HD IVS',
    'Arricam ST Camera',
    'Arricam ST HD IVS',
    'Arri 416 Plus Camera',
    'Indieassist 416 HD IVS ',
    'X = Global Shutter Sensor' // Raw value that gets translated
  ];

  // List of valid todayCell background colors
  const validTodayCellBackgrounds = [
    '#ffffff', // white
    '#f9ff71', // yellow
    '#66ff75', // green
    '#4a86e8', // blue
    '#ff7171', // red
    '#00ffff'  // cyan
  ];

  Logger.log('Processing camera scheduling data...');
  let processedCameras = 0;
  let validCameras = 0;
  
  // Pre-filter valid cameras to avoid repeated validation
  const validCamerasToProcess = cameraTypes.filter(camera => validCameraTypes.includes(camera.type));
  validCameras = validCamerasToProcess.length;
  processedCameras = cameraTypes.length;
  
  // Batch get all data at once for better performance
  const startCol = todayColumn;
  const numCols = 8; // today + next 7
  
  // Get all camera rows data in one batch operation
  const cameraRows = validCamerasToProcess.map(camera => camera.row);
  const minRow = Math.min(...cameraRows);
  const maxRow = Math.max(...cameraRows);
  const rowRange = maxRow - minRow + 1;
  
  // Get all values, backgrounds, and date headers in batch operations
  const allValues = cameraSheet.getRange(minRow, startCol, rowRange, numCols).getValues();
  const allBackgrounds = cameraSheet.getRange(minRow, startCol, rowRange, numCols).getBackgrounds();
  const dateHeaders = cameraSheet.getRange(1, startCol, 1, numCols).getValues()[0];
  
  // Create a map for quick row access
  const rowDataMap = {};
  for (let i = 0; i < cameraRows.length; i++) {
    const rowIndex = cameraRows[i] - minRow;
    rowDataMap[cameraRows[i]] = {
      values: allValues[rowIndex],
      backgrounds: allBackgrounds[rowIndex]
    };
  }
  
  // Process each valid camera
  for (const camera of validCamerasToProcess) {
          // Extract barcode from column E of the found row
      const barcodeCell = data[camera.row - 1][4]; // Column E, 0-based index
      let cameraBarcode = '';
      if (typeof barcodeCell === 'string') {
        // Updated regex to handle alphanumeric barcodes with hyphens (e.g., BC#ALX35-3)
        const match = barcodeCell.match(/BC#\s*([A-Z0-9-]+)/);
        if (match) {
          cameraBarcode = match[1];
        }
      }

    // Get cached data for this camera
    const rowData = rowDataMap[camera.row];
    const values = rowData.values;
    const backgrounds = rowData.backgrounds;

    // Only scan next 7 days if today cell background is in the valid list
    if (validTodayCellBackgrounds.includes(backgrounds[0])) {
      let foundColor = null;
      for (let i = 1; i < numCols; i++) { // skip today (i=0), only look at Day +1 to Day +7
        const val = values[i];
        if (
          val !== '' &&
          typeof val === 'string' &&
          !val.toLowerCase().includes('reserve') &&
          !val.toLowerCase().includes('repair') &&
          !val.toLowerCase().includes('in progress') &&
          !val.toLowerCase().includes('rtr')
        ) {
          // Check if this is an out-of-town job for an LA camera
          const isOutOfTownJob = hasOutOfTownPrefix(val);
          let shouldSkipCamera = false;
          
          if (isOutOfTownJob) {
            Logger.log(`Out-of-town job detected for LA camera "${cameraBarcode}": "${val}"`);
            
            // Get the full row data for this camera to check for return to LA
            const fullRowValues = data[camera.row - 1];
            const fullRowBackgrounds = cameraSheet.getRange(camera.row, 1, 1, fullRowValues.length).getBackgrounds()[0];
            
            // Check if the colored segment contains "returns to LA" or similar
            const hasReturnToLA = checkForReturnToLA(fullRowValues, fullRowBackgrounds, startCol + i - 1, fullRowValues.length);
            
            if (hasReturnToLA) {
              shouldSkipCamera = true;
              Logger.log(`Camera "${cameraBarcode}" EXCLUDED from forecast - out-of-town job with return to LA (will leave LA): "${val}"`);
            } else {
              Logger.log(`Camera "${cameraBarcode}" INCLUDED in forecast - out-of-town job without return to LA (stays in LA): "${val}"`);
            }
          }
          
          // Skip this camera if it's an out-of-town job with return to LA
          if (shouldSkipCamera) {
            break;
          }
          
          // Set foundColor to the background of the first valid non-empty cell
          if (foundColor === null) {
            foundColor = backgrounds[i];
          }
          // Start a new 7-day search from the first valid cell
          for (let j = i + 1; j <= i + 7 && j < numCols; j++) {
            // Check for color change regardless of cell content
              if (backgrounds[j] !== foundColor && backgrounds[j] !== '#ffffff') {
                foundColor = backgrounds[j];
                break; // Stop after the first color change
            }
          }
          // If no color change is found, foundColor remains as is
          // Weekend offset logic
          let foundDate = new Date(dateHeaders[i]);
          let offsetDate = new Date(foundDate);
          let offsetApplied = false;
          if (foundDate.getDay() === 0) { // Sunday
            offsetDate.setDate(foundDate.getDate() - 1); // Move to Saturday
            offsetApplied = true;
          } else if (foundDate.getDay() === 1) { // Monday
            offsetDate.setDate(foundDate.getDate() - 2); // Move to Saturday
            offsetApplied = true;
          }
          const offsetDateStr = `${offsetDate.getMonth() + 1}/${offsetDate.getDate()}/${offsetDate.getFullYear()}`;
          Logger.log(`Regular forecast: Adding camera Type="${camera.type}" Barcode="${cameraBarcode}" TodayPlus="Today +${i}" Job="${val}" Date="${offsetApplied ? offsetDateStr : dateHeaders[i]}"`);
          outputRows.push([
            camera.type,
            cameraBarcode,
            `Today +${i}`,
            val,
            offsetApplied ? offsetDateStr : dateHeaders[i]
          ]);
          // Store foundColor in the map using barcode as the key
          rowColorMap[cameraBarcode] = foundColor;
          outputBackgrounds.push(['', '', '', foundColor, '']);
          break; // Only the first valid non-empty cell is considered
        }
      }
    }
  }
  Logger.log(`Camera processing complete: ${processedCameras} total cameras, ${validCameras} valid cameras, ${outputRows.length} job assignments found`);

  // --------------------- PROCESS ORDER NUMBERS FROM TOMORROW DOUBLE CHECK ---------------------
  try {
    const forecastSheetTmp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LA Camera Forecast');
    if (forecastSheetTmp) {
      const orderRange = forecastSheetTmp.getRange(2, 37, forecastSheetTmp.getLastRow() - 1, 1);
      const orderValues = orderRange.getValues().flat().map(v => v && v.toString().trim()).filter(v => v);
      const orderSet = new Set(orderValues);
      Logger.log(`Processing ${orderSet.size} order numbers from AJ:AM section...`);

      const bgCache = {};
      const sheetLastCol = headerRow.length;

      orderSet.forEach(ord => {
        Logger.log(`Searching for order ${ord} across entire scheduling sheet...`);
        let camerasFoundForOrder = 0;
        let totalMatches = 0;
        let skippedMatches = 0;

        // Search every row in the data for the order number
        const orderRegex = new RegExp("\\b" + ord + "\\b");
        for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
          const rowNum = rowIdx + 1; // Convert to 1-based
          const rowVals = data[rowIdx];
          
          // First check if the order number exists in columns F and beyond
          let hasOrderNumber = false;
          for (let colIdx = 5; colIdx < rowVals.length; colIdx++) {
            const cellValue = rowVals[colIdx];
            if (cellValue && orderRegex.test(cellValue.toString())) {
              hasOrderNumber = true;
              totalMatches++;
              break;
            }
          }
          
          // If we found the order number, check if this is a valid camera row
          if (hasOrderNumber) {
            const isLACamera = rowVals[0] === 'LOS ANGELES';
            const isGTCamera = gtCheck(rowVals);
            
            if (!isLACamera && !isGTCamera) {
              skippedMatches++;
              Logger.log(`  Found order ${ord} at row ${rowNum} but SKIPPED - Column A: "${rowVals[0]}" (not LA camera or valid GT)`);
              continue;
            }
          } else {
            continue; // No order number in this row, skip entirely
          }
          
          // Look for order number starting from column F (index 5) and beyond
          for (let colIdx = 5; colIdx < rowVals.length; colIdx++) {
            const cellValue = rowVals[colIdx];
            if (cellValue && orderRegex.test(cellValue.toString())) {
              Logger.log(`  Match: Order ${ord} at row ${rowNum}, col ${colIdx+1} in cell: "${cellValue}"`);
              
              // Get backgrounds for this row if not cached
              if (!bgCache[rowNum]) {
                bgCache[rowNum] = cameraSheet.getRange(rowNum, 1, 1, sheetLastCol).getBackgrounds()[0];
              }
              const bgRow = bgCache[rowNum];
              const cellBg = bgRow[colIdx];
              
              Logger.log(`    Cell background: ${cellBg}`);
              if (!validTodayCellBackgrounds.includes(cellBg)) {
                Logger.log('    Skipped due to invalid background color');
                continue;
              }
              
              // Find camera type by looking for the first empty cell above
              let cameraType = '';
              let typeRow = rowIdx - 1;
              while (typeRow >= 0 && data[typeRow][0] !== '') {
                typeRow--;
              }
              if (typeRow >= 0) {
                const rawCameraType = data[typeRow][4]; // Column E
                cameraType = translateCameraType(rawCameraType); // Apply translation rules
              }
              
              // Extract barcode from column E
              const barcodeCell = rowVals[4];
              let cameraBarcode = '';
              if (typeof barcodeCell === 'string') {
                const m = barcodeCell.match(/BC#\s*([A-Z0-9-]+)/);
                if (m) cameraBarcode = m[1];
              }
              
              const headerCellVal = headerRow[colIdx];
              const dateStr = headerCellVal instanceof Date ? formatDate(headerCellVal) : headerCellVal;
              
              // Check if this is an out-of-town job for an LA camera
              const isOutOfTownJob = hasOutOfTownPrefix(cellValue);
              let shouldSkipCamera = false;
              
              if (isOutOfTownJob) {
                Logger.log(`    Out-of-town job detected for LA camera "${cameraBarcode}": "${cellValue}"`);
                
                // Check if the colored segment contains "returns to LA" or similar
                const hasReturnToLA = checkForReturnToLA(rowVals, bgRow, colIdx, rowVals.length);
                
                if (hasReturnToLA) {
                  shouldSkipCamera = true;
                  Logger.log(`    Camera "${cameraBarcode}" EXCLUDED from forecast - out-of-town job with return to LA (will leave LA): "${cellValue}"`);
                } else {
                  Logger.log(`    Camera "${cameraBarcode}" INCLUDED in forecast - out-of-town job without return to LA (stays in LA): "${cellValue}"`);
                }
              }
              
              // Skip this camera if it's an out-of-town job with return to LA
              if (shouldSkipCamera) {
                break;
              }
              
              // Look for color change in next 7 days from this match
              let foundColor = cellBg;
              for (let j = colIdx + 1; j <= colIdx + 7 && j < rowVals.length; j++) {
                const nextCellBg = bgRow[j];
                if (nextCellBg !== foundColor && nextCellBg !== '#ffffff') {
                  foundColor = nextCellBg;
                  break;
                }
              }
              
              Logger.log(`    Adding camera to output: Type="${cameraType}" Barcode="${cameraBarcode}" Job="${cellValue}" Date="${dateStr}" (found at row ${rowNum}, col ${colIdx+1})`);
              outputRows.push([cameraType, cameraBarcode, 'Today +0', cellValue, dateStr]);
              outputBackgrounds.push(['', '', '', foundColor, '']);
              rowColorMap[cameraBarcode] = foundColor;
              camerasFoundForOrder++;
              break; // Still break from column loop, but continue with next row
            }
          }
          // Removed the foundForOrder break here - continue searching all rows
        }

        Logger.log(`  Order ${ord} search complete: ${totalMatches} total matches found, ${skippedMatches} skipped (wrong location/GT), ${camerasFoundForOrder} cameras added to forecast.`);
        if (camerasFoundForOrder === 0) {
          Logger.log(`  Order ${ord} NOT FOUND in valid camera rows (LA or valid GT).`);
        } else {
          Logger.log(`  Order ${ord} FOUND: ${camerasFoundForOrder} cameras added to forecast.`);
        }
      });
    }
  } catch (err) {
    Logger.log('Error processing tomorrow order numbers: ' + err);
  }



  // Print today's date and the next 7 days' dates to B2:I2 in the same format as in AE:AE
  const forecastSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LA Camera Forecast');
  if (forecastSheet) {
    const dateRow = [];
    const bgRow = [];
    for (let offset = 0; offset < 8; offset++) {
      const d = new Date(today.getTime());
      d.setDate(today.getDate() + offset);
      dateRow.push(formatDate(d));
      // Set background to Dark Magenta 1 (#6a329f) for weekends, Dark Green 1 (#62a162) for weekdays
      if (d.getDay() === 0 || d.getDay() === 6) {
        bgRow.push('#6a329f');
      } else {
        bgRow.push('#62a162');
      }
    }
    forecastSheet.getRange(2, 2, 1, 8).setValues([dateRow]);
    forecastSheet.getRange(2, 2, 1, 8).setBackgrounds([bgRow]);
  }

  let sortedRows = [];
  let sortedBackgrounds = [];

  // Sort outputRows and outputBackgrounds by Today +x in ascending order
  if (outputRows.length > 0) {
    // Extract the numeric part from 'Today +x' for sorting
    const getDayNumber = row => {
      const match = row[2].match(/Today \+(\d+)/);
      return match ? parseInt(match[1], 10) : 0;
    };
    // Combine rows and backgrounds for stable sorting
    const combined = outputRows.map((row, idx) => ({ row, bg: outputBackgrounds[idx] }));
    combined.sort((a, b) => getDayNumber(a.row) - getDayNumber(b.row));
    // Unpack sorted arrays
    sortedRows = combined.map(item => item.row);
    sortedBackgrounds = combined.map(item => item.bg);

          // Log summary before writing to sheet
      Logger.log(`Preparing to write ${sortedRows.length} rows to forecast sheet`);

    if (forecastSheet) {
      // Clear all data in AA:AF starting from row 2 (now 6 columns)
      const lastRow = forecastSheet.getLastRow();
      if (lastRow > 1) {
        forecastSheet.getRange(2, 27, lastRow - 1, 6).clearContent().clearFormat();
      }
      // RTR Status logic for column AF - MOVED TO AFTER FINAL SORTING
      // Update the date for each camera/job if found in Prep Bay Assignments BEFORE writing to sheet
      Logger.log('Updating dates from prep bay assignments...');
      let dateUpdates = 0;
      for (let i = 0; i < sortedRows.length; i++) {
        const row = sortedRows[i];
        const jobString = row[3]; // Assuming job string is in the 4th column of sortedRows
        const { orderNumber, jobName } = extractOrderAndJobNameFromSchedule(jobString);
        const normalizedJobName = normalizeString(jobName);

        // Fetch data from the prep bay assignment map
        const prepBayEntry = Object.entries(jobDateMap).find(([key, value]) => 
          (normalizeString(key.replace(/"/g, '')) === normalizedJobName || String(value.orderNumber).trim() === orderNumber.trim())
        );

        if (prepBayEntry) {
          // Update the date in the outputRows
          const prepBayDate = prepBayEntry[1].date;
          if (prepBayDate && prepBayDate !== '') {
            // If the date is already in MM/DD/YYYY format, use it directly
            if (prepBayDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
              row[4] = prepBayDate;
            } else {
              // If it's in MM/DD format, add the current year
              const dateMatch = prepBayDate.match(/^(\d{1,2})\/(\d{1,2})$/);
              if (dateMatch) {
                const month = parseInt(dateMatch[1], 10) - 1; // JavaScript months are 0-based
                const day = parseInt(dateMatch[2], 10);
                const dateWithYear = new Date(today.getFullYear(), month, day);
                
                // Adjust year if the date is in the past
                if (dateWithYear < today) {
                  dateWithYear.setFullYear(today.getFullYear() + 1);
                }
                row[4] = formatDate(dateWithYear);
              } else {
                row[4] = prepBayDate; // Use as-is if no pattern matches
              }
            }
            dateUpdates++;
          }
        }
      }
      Logger.log(`Date updates complete: ${dateUpdates} dates updated from prep bay assignments`);

      // Process notes for 'crew prep' and update applicable date BEFORE writing to sheet
      let crewPrepUpdates = 0;
      sortedRows.forEach((row, index) => {
        const note = row[3]; // Assuming note is in the 4th column of sortedRows
        // Convert to string and handle null/undefined values
        const noteString = (note !== null && note !== undefined) ? note.toString() : '';
        if (noteString && noteString.toLowerCase().includes('crew prep')) {
          const dateMatches = noteString.match(/\d{1,2}\/\d{1,2}/g); // Match dates like 6/5 or 06/05
          if (dateMatches) {
            const earliestDate = dateMatches.reduce((earliest, current) => {
              const [month, day] = current.split('/').map(Number);
              const currentDate = new Date(today.getFullYear(), month - 1, day);
              return currentDate < earliest ? currentDate : earliest;
            }, new Date(today.getFullYear() + 1, 0, 1)); // Start with a future date for comparison
            row[4] = formatDate(earliestDate); // Update the applicable date
            crewPrepUpdates++;
          }
        }
      });
      if (crewPrepUpdates > 0) {
        Logger.log(`Crew prep date updates: ${crewPrepUpdates} dates updated from crew prep notes`);
      }

      // Function to calculate forecast day offset ensuring we never forecast FOR a weekend
      function calculateForecastDay(fromDate, toDate) {
        const from = new Date(fromDate);
        const to = new Date(toDate);
        
        // Calculate business days between dates
        let businessDays = 0;
        const currentDate = new Date(from);
        currentDate.setDate(currentDate.getDate() + 1);
        
        while (currentDate <= to) {
          const dayOfWeek = currentDate.getDay();
          if (dayOfWeek >= 1 && dayOfWeek <= 5) { // Monday through Friday
            businessDays++;
          }
          currentDate.setDate(currentDate.getDate() + 1);
        }
        
        // Rule: We can never forecast FOR a weekend
        // Calculate what day "Today +businessDays" would be
        const todayDayOfWeek = from.getDay();
        const forecastDay = new Date(from);
        forecastDay.setDate(forecastDay.getDate() + businessDays);
        const forecastDayOfWeek = forecastDay.getDay();
        
        let adjustedOffset = businessDays;
        
        // If Today +X would be a weekend, adjust it to the previous business day
        if (forecastDayOfWeek === 6) { // Saturday
          // Move back to Friday
          adjustedOffset = Math.max(0, businessDays - 1);
          Logger.log(`  Weekend adjustment: Today +${businessDays} would be Saturday, adjusted to Today +${adjustedOffset}`);
        } else if (forecastDayOfWeek === 0) { // Sunday
          // Move back to Friday
          adjustedOffset = Math.max(0, businessDays - 2);
          Logger.log(`  Weekend adjustment: Today +${businessDays} would be Sunday, adjusted to Today +${adjustedOffset}`);
        }
        
        // Additional rule: if today is Friday, Today +1 and Today +2 don't exist (would be weekend)
        if (todayDayOfWeek === 5) { // Friday
          if (adjustedOffset === 1 || adjustedOffset === 2) {
            adjustedOffset = 0; // Move to Today
            Logger.log(`  Friday special rule: Today +1 and +2 not allowed, changed to Today`);
          }
        }
        
        // Additional rule: if today is Thursday, Today +2 doesn't exist (would be Saturday)
        if (todayDayOfWeek === 4) { // Thursday
          if (adjustedOffset === 2) {
            adjustedOffset = 1; // Move to Today +1 (Friday)
            Logger.log(`  Thursday special rule: Today +2 would be Saturday, changed to Today +1`);
          }
        }
        
        return Math.max(0, adjustedOffset);
      }

      // Recalculate "Today +" values using business days after date updates
      Logger.log('Recalculating "Today +" values using business days after date updates...');
      let recalculatedCount = 0;
      for (let i = 0; i < sortedRows.length; i++) {
        const row = sortedRows[i];
        const updatedDate = row[4];
        
        if (updatedDate && updatedDate !== '') {
          // Parse the updated date
          let parsedDate;
          if (typeof updatedDate === 'string') {
            // Handle MM/DD/YYYY format
            const dateMatch = updatedDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
            if (dateMatch) {
              const month = parseInt(dateMatch[1], 10) - 1; // JavaScript months are 0-based
              const day = parseInt(dateMatch[2], 10);
              const year = parseInt(dateMatch[3], 10);
              parsedDate = new Date(year, month, day);
            }
          } else if (updatedDate instanceof Date) {
            parsedDate = updatedDate;
          }
          
                    if (parsedDate) {
            const crewPrepDate = new Date(parsedDate);
            
            // Calculate the service date (day before crew prep, but service cannot be on weekend)
            const serviceDate = new Date(crewPrepDate);
            serviceDate.setDate(serviceDate.getDate() - 1); // Start with day before
            
            // If service date would be a weekend, move it back to Friday
            const serviceDayOfWeek = serviceDate.getDay();
            if (serviceDayOfWeek === 0) { // Sunday → move to Friday
              serviceDate.setDate(serviceDate.getDate() - 2);
            } else if (serviceDayOfWeek === 6) { // Saturday → move to Friday
              serviceDate.setDate(serviceDate.getDate() - 1);
            }
            
            Logger.log(`  Service date calc: CrewPrep=${formatDate(crewPrepDate)} → Service=${formatDate(serviceDate)}`);
            
            // Calculate ALL calendar days from today to service date
            const timeDiff = serviceDate.getTime() - today.getTime();
            const calendarDays = Math.ceil(timeDiff / (1000 * 3600 * 24));
            
            Logger.log(`  Calendar days to service: ${calendarDays}`);
            
            // Update the "Today +" value based on calendar days to service date
            if (calendarDays <= 0) {
              row[2] = 'Today';
            } else {
              row[2] = `Today +${calendarDays}`;
            }
            recalculatedCount++;
          }
        }
      }
      Logger.log(`Recalculated "Today +" values for ${recalculatedCount} rows`);

      // Re-sort rows based on updated "Today +" values
      Logger.log('Re-sorting rows based on updated "Today +" values...');
      const getDayNumber = row => {
        const match = row[2].match(/Today \+(\d+)/);
        return match ? parseInt(match[1], 10) : 0;
      };
      sortedRows.sort((a, b) => getDayNumber(a) - getDayNumber(b));
      Logger.log(`Re-sorting complete: ${sortedRows.length} rows sorted by updated "Today +" values`);

      // RTR Status logic for column AF - PROCESS AFTER FINAL SORTING
      // Camera type to sheet mapping
      const cameraTypeToSheet = {
        'ARRI ALEXA 35 Camera Body': 'ALEXA 35 Body Status',
        'ARRI ALEXA Mini LF Camera Body': 'Alexa Mini LF Body Status',
        'ARRI ALEXA Mini Camera Body': 'Alexa Mini Body Status',
        'SONY VENICE 2': 'VENICE 2 Body Status',
        'SONY VENICE 1': 'Venice 1 Body Status',
        "Sony BURANO Digital Camera" : "BURANO STATUS",
        "Sony FX3 Digital Camera" : "FX3 STATUS",
        "Sony FX6 Digital Camera" : "FX6 STATUS",
        "Sony PXW-FX9 Digital Camera" : "FX9 STATUS",
        "RED V-RAPTOR DSMC3 [X] 8K VV Digital Camera" : "V-RAPTOR [X] STATUS",
        "RED V-RAPTOR XL [X] 8K VV Digital Camera" : "V-RAPTOR XL [X] STATUS"
      };
      // Initialize rtrStatusResults array
      const rtrStatusResults = [];
      // Initialize background colors array
      const rtrStatusBackgrounds = [];
      // Preload all database sheets into a map for efficiency
      const dbSheetsData = {};
      Logger.log('Loading database sheets...');
      
      // Debug: List all available sheets in the spreadsheet
      const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      Logger.log('Available sheets in spreadsheet:');
      allSheets.forEach(sheet => Logger.log(`  - "${sheet.getName()}"`));
      
      for (const sheetName of Object.values(cameraTypeToSheet)) {
        const dbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        Logger.log(`Looking for sheet: "${sheetName}" - Found: ${dbSheet ? 'YES' : 'NO'}`);
        dbSheetsData[sheetName] = dbSheet ? dbSheet.getDataRange().getValues() : null;
      }
      Logger.log('Processing RTR status for cameras...');
      for (let i = 0; i < sortedRows.length; i++) {
        const row = sortedRows[i];
        const cameraType = row[0];
        const barcode = row[1]; // Use barcode from column 1 (before swap)
        let statusResult = '';
        let statusBackground = '#ffffff'; // Default white background
        const dbSheetName = cameraTypeToSheet[cameraType];
        Logger.log(`Processing camera: "${cameraType}" with barcode: "${barcode}" - Looking for sheet: "${dbSheetName}"`);
        if (!dbSheetName) {
          statusResult = 'No Database Found';
        } else {
          const dbData = dbSheetsData[dbSheetName];
          if (!dbData) {
            statusResult = 'No Database Found';
          } else {
            let foundBarcode = false;
            for (let r = 0; r < dbData.length; r++) {
              if (dbData[r][3] && dbData[r][3].toString() === barcode) { // Column D (index 3)
                foundBarcode = true;
                const status = dbData[r][4] ? dbData[r][4].toString() : '';
                if (status === 'Left Inventory' || status === 'Disposed') {
                  statusResult = '';
                } else {
                  statusResult = status || 'No Service Status Found';
                  // Set green background for RTR or Serviced status
                  if (status === 'RTR' || status === 'Serviced') {
                    statusBackground = '#66ff75'; // Green background
                  }
                }
                break;
              }
            }
            if (!foundBarcode) statusResult = 'No Matching Camera Found';
          }
        }
        rtrStatusResults.push([statusResult]);
        rtrStatusBackgrounds.push([statusBackground]);
      }
      Logger.log(`RTR status processing complete: ${rtrStatusResults.length} cameras processed`);

      // ------------------ Remove duplicate barcode+order pairs BEFORE column swap ------------------
      Logger.log(`Starting deduplication process with ${sortedRows.length} rows`);
      if (sortedRows.length > 0) {
        const seenPairs = new Set();
        const dedupRows = [];
        const dedupBgs  = [];
        let duplicatesRemoved = 0;
        let keptWithMissingData = 0;
        
        for (let i = 0; i < sortedRows.length; i++) {
          const row = sortedRows[i];
          const barcode = row[1];  // Barcode is still in position 1 before swap
          const jobText = row[3];  // Job text in position 3
          const cameraType = row[0];
          const todayPlus = row[2];
          
          // Extract order number from job text using regex (matches Google Sheets REGEXEXTRACT)
          let orderFromJob = '';
          const orderMatch = jobText && jobText.toString().match(/\d{6}/);
          if (orderMatch) {
            orderFromJob = orderMatch[0];
          }
          
          const pairKey = `${barcode}-${orderFromJob}`;
          
          // Log every row being processed
          Logger.log(`Dedup Row ${i+1}: Camera="${cameraType}" Barcode="${barcode}" Order="${orderFromJob}" Job="${jobText}" TodayPlus="${todayPlus}"`);
          
          // Only remove duplicates if both barcode and order number are valid (not empty/null)
          if (barcode && orderFromJob && !seenPairs.has(pairKey)) {
            seenPairs.add(pairKey);
            dedupRows.push(row);
            dedupBgs.push(sortedBackgrounds[i]);
            Logger.log(`  → KEPT: First occurrence of pair ${pairKey}`);
          } else if (barcode && orderFromJob && seenPairs.has(pairKey)) {
            duplicatesRemoved++;
            Logger.log(`  → REMOVED: Duplicate barcode+order pair: ${barcode} + ${orderFromJob} (${cameraType})`);
          } else {
            // Keep rows where barcode or order is missing/empty (no deduplication)
            dedupRows.push(row);
            dedupBgs.push(sortedBackgrounds[i]);
            keptWithMissingData++;
            Logger.log(`  → KEPT: Missing data (barcode="${barcode}" order="${orderFromJob}")`);
          }
        }
        sortedRows.length = 0;
        sortedBackgrounds.length = 0;
        sortedRows.push(...dedupRows);
        sortedBackgrounds.push(...dedupBgs);
        Logger.log(`Deduplication complete: ${duplicatesRemoved} duplicates removed, ${keptWithMissingData} kept with missing data, ${sortedRows.length} total rows remain.`);
      }

      // Build AE background color array from rowColorMap using barcode (for job info column)
      const aeBackgrounds = sortedRows.map(row => [rowColorMap[row[1]] || '#ffffff']);

      // Swap barcode and date columns before writing to sheet
      Logger.log('Swapping barcode and date columns...');
      for (let i = 0; i < sortedRows.length; i++) {
        const row = sortedRows[i];
        const temp = row[1]; // Store barcode
        row[1] = row[4];     // Move date to barcode position
        row[4] = temp;       // Move barcode to date position
      }

      // --------------------- ADJUST CAMERAS FOR TOMORROW PREP ORDERS ---------------------
      // Read the tomorrow prep order numbers and adjust any matching cameras to "Today"
      try {
        const tomorrowPrepOrderRange = forecastSheet.getRange(2, 37, forecastSheet.getLastRow() - 1, 1); // Column AK
        const tomorrowPrepOrders = tomorrowPrepOrderRange.getValues().flat()
          .map(v => v && v.toString().trim())
          .filter(v => v)
          .map(v => v.replace(/[^0-9]/g, '')); // Normalize to digits only
        
        const tomorrowPrepOrderSet = new Set(tomorrowPrepOrders);
        Logger.log(`Adjusting cameras for ${tomorrowPrepOrderSet.size} tomorrow prep orders...`);
        
        let adjustedCameras = 0;
        for (let i = 0; i < sortedRows.length; i++) {
          const row = sortedRows[i];
          const jobText = row[3]; // Job text column
          
          // Extract order number from job text
          const jobString = (jobText !== null && jobText !== undefined) ? jobText.toString() : '';
          const orderMatch = jobString.match(/\b\d{6}\b/);
          const orderNumber = orderMatch ? orderMatch[0] : '';
          
          // If this camera's order is in the tomorrow prep list, adjust to "Today"
          if (orderNumber && tomorrowPrepOrderSet.has(orderNumber)) {
            const oldTodayPlus = row[2];
            row[2] = 'Today';
            adjustedCameras++;
            Logger.log(`  Adjusted camera: Order ${orderNumber} changed from "${oldTodayPlus}" to "Today" (tomorrow prep confirmed)`);
          }
        }
        Logger.log(`Tomorrow prep adjustment complete: ${adjustedCameras} cameras adjusted to "Today"`);
      } catch (err) {
        Logger.log('Error adjusting cameras for tomorrow prep: ' + err);
      }

      // Log final results before writing to sheet
      Logger.log(`Final forecast results (${sortedRows.length} cameras):`);
      sortedRows.forEach((row, idx) => {
        // Note: After column swap, row[4] contains barcode and row[1] contains date
        Logger.log(`  Final Row ${idx+1}: Type="${row[0]}" Date="${row[1]}" TodayPlus="${row[2]}" Job="${row[3]}" Barcode="${row[4]}"`);
      });

      // Write the rest of the output starting at AB2 (col 28)
      forecastSheet.getRange(2, 28, sortedRows.length, sortedRows[0].length).setValues(sortedRows);
      // Set background color for column AE (col 31, which is the job info column)
      forecastSheet.getRange(2, 31, aeBackgrounds.length, 1).setBackgrounds(aeBackgrounds);
      
      // Write RTR status results to column AA (col 27)
      forecastSheet.getRange(2, 27, rtrStatusResults.length, 1).setValues(rtrStatusResults);
      // Set background colors for RTR status column
      forecastSheet.getRange(2, 27, rtrStatusBackgrounds.length, 1).setBackgrounds(rtrStatusBackgrounds);

      // Check for duplicate barcodes in AF2:AF and assign 'TurnAroundColor'
      const barcodeRange = forecastSheet.getRange(2, 32, sortedRows.length, 1); // Column AF is index 32
      const barcodes = barcodeRange.getValues().flat();
      const barcodeCountMap = {};
      barcodes.forEach(barcode => {
        if (barcode && barcode.toString().trim() !== '') {
          barcodeCountMap[barcode] = (barcodeCountMap[barcode] || 0) + 1;
        }
      });

      barcodes.forEach((barcode, index) => {
        if (barcode && barcode.toString().trim() !== '' && barcodeCountMap[barcode] > 1) {
          barcodeRange.getCell(index + 1, 1).setBackground('#f4e198'); // Assign 'TurnAroundColor'
        }
      });

      Logger.log(`Wrote ${sortedRows.length} rows to 'LA Camera Forecast' sheet starting at AB2, with RTR status in AA.`);

      // AJ:AM already refreshed at start of getCameraForecast()
    } else {
      Logger.log("Sheet 'LA Camera Forecast' not found in the active spreadsheet.");
    }
  } else {
    Logger.log('No results to write to LA Camera Forecast sheet.');
  }

  // Duplicate removal now handled before column swap above


}

/**
 * prepBayBlock
 * Scans Prep Bay Assignment sheets for today, tomorrow, and day after tomorrow
 * and writes them (with date header) to columns AI:AM of the 'LA Camera Forecast' sheet.
 * Column AI1 contains today's date in "Thurs 11/6" format.
 */
function prepBayBlock() {
  // --- Date helpers ---
  const today = new Date();
  const dayPrefixes = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];
  const expectedSheetName = (dateObj) => `${dayPrefixes[dateObj.getDay()]} ${dateObj.getMonth()+1}/${dateObj.getDate()}`;
  
  // Calculate next business day (skip weekends)
  function getNextBusinessDay(date) {
    const nextDay = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayOfWeek = date.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
    
    if (dayOfWeek === 5) { // Friday
      nextDay.setDate(date.getDate() + 3); // Skip to Monday
    } else if (dayOfWeek === 6) { // Saturday
      nextDay.setDate(date.getDate() + 2); // Skip to Monday
    } else if (dayOfWeek === 0) { // Sunday
      nextDay.setDate(date.getDate() + 1); // Skip to Monday
    } else { // Monday-Thursday
      nextDay.setDate(date.getDate() + 1); // Next day
    }
    
    return nextDay;
  }
  
  // Calculate day after tomorrow (2 business days from today)
  function getDayAfterTomorrow(date) {
    const tomorrow = getNextBusinessDay(date);
    return getNextBusinessDay(tomorrow);
  }
  
  const tomorrow = getNextBusinessDay(today);
  const dayAfterTomorrow = getDayAfterTomorrow(today);

  // --- Small helpers ---
  const normalizeString = (str) => String(str || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const normalizeOrder = (val) => String(val || '').replace(/[^0-9]/g, '');

  // --- Open Prep Bay workbook ---
  const prepBaySS = SpreadsheetApp.openById('1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ');
  const visibleSheets = prepBaySS.getSheets().filter(sh => !sh.isSheetHidden());

  // Locate today, tomorrow, and day after tomorrow sheets
  const todaySheetName = expectedSheetName(today);
  const tomorrowSheetName = expectedSheetName(tomorrow);
  const dayAfterTomorrowSheetName = expectedSheetName(dayAfterTomorrow);
  
  const todaySheet = visibleSheets.find(sh => sh.getName() === todaySheetName);
  const tomorrowSheet = visibleSheets.find(sh => sh.getName() === tomorrowSheetName);
  const dayAfterTomorrowSheet = visibleSheets.find(sh => sh.getName() === dayAfterTomorrowSheetName);

  // Collect jobs for each day
  const todayJobs = [];
  const tomorrowJobs = [];
  const dayAfterTomorrowJobs = [];
  const addedEntries = new Set(); // Track complete entries to prevent duplicates

  // Process today's sheet
  if (todaySheet) {
    const rows = todaySheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName = r[0];
      const orderNum = normalizeOrder(r[1]);
      const cameraInfo = r[3]; // Column E (was F before column D deletion)
      const note = r[7]; // Column I (was J before column D deletion)
      
      if (!jobName && !orderNum) return; // blank row
      
      const lowerName = normalizeString(jobName);
      const entryKey = `${lowerName}-${orderNum}-${normalizeString(cameraInfo)}`;
      
      // Skip wrap out jobs and duplicates
      if (!lowerName.includes('wrap out') && !addedEntries.has(entryKey)) {
        todayJobs.push([jobName, orderNum, cameraInfo, note]);
        addedEntries.add(entryKey);
      }
    });
    Logger.log(`Prep Bay Block: collected ${todayJobs.length} jobs from today (${todaySheetName}).`);
  } else {
    Logger.log(`Prep Bay Block: today's sheet "${todaySheetName}" not found.`);
  }

  // Process tomorrow's sheet
  if (tomorrowSheet) {
    const rows = tomorrowSheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName = r[0];
      const orderNum = normalizeOrder(r[1]);
      const cameraInfo = r[3]; // Column E (was F before column D deletion)
      const note = r[7]; // Column I (was J before column D deletion)
      
      if (!jobName && !orderNum) return; // blank row
      
      const lowerName = normalizeString(jobName);
      const entryKey = `${lowerName}-${orderNum}-${normalizeString(cameraInfo)}`;
      
      // Skip wrap out jobs and duplicates
      if (!lowerName.includes('wrap out') && !addedEntries.has(entryKey)) {
        tomorrowJobs.push([jobName, orderNum, cameraInfo, note]);
        addedEntries.add(entryKey);
      }
    });
    Logger.log(`Prep Bay Block: collected ${tomorrowJobs.length} jobs from tomorrow (${tomorrowSheetName}).`);
  } else {
    Logger.log(`Prep Bay Block: tomorrow's sheet "${tomorrowSheetName}" not found.`);
  }

  // Process day after tomorrow's sheet
  if (dayAfterTomorrowSheet) {
    const rows = dayAfterTomorrowSheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName = r[0];
      const orderNum = normalizeOrder(r[1]);
      const cameraInfo = r[3]; // Column E (was F before column D deletion)
      const note = r[7]; // Column I (was J before column D deletion)
      
      if (!jobName && !orderNum) return; // blank row
      
      const lowerName = normalizeString(jobName);
      const entryKey = `${lowerName}-${orderNum}-${normalizeString(cameraInfo)}`;
      
      // Skip wrap out jobs and duplicates
      if (!lowerName.includes('wrap out') && !addedEntries.has(entryKey)) {
        dayAfterTomorrowJobs.push([jobName, orderNum, cameraInfo, note]);
        addedEntries.add(entryKey);
      }
    });
    Logger.log(`Prep Bay Block: collected ${dayAfterTomorrowJobs.length} jobs from day after tomorrow (${dayAfterTomorrowSheetName}).`);
  } else {
    Logger.log(`Prep Bay Block: day after tomorrow's sheet "${dayAfterTomorrowSheetName}" not found.`);
  }

  // --- Write to LA Camera Forecast sheet (AI:AM) ---
  const forecastSpreadsheetId = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
  const forecastSpreadsheet = SpreadsheetApp.openById(forecastSpreadsheetId);
  const forecastSheet = forecastSpreadsheet.getSheetByName('LA Camera Forecast');
  
  if (!forecastSheet) {
    Logger.log('Prep Bay Block: LA Camera Forecast sheet not found.');
    return;
  }

  // Write date header to AI1 (row 1, column 35)
  const dateHeader = `${dayPrefixes[today.getDay()]} ${today.getMonth() + 1}/${today.getDate()}`;
  forecastSheet.getRange(1, 35, 1, 1).setValue(dateHeader).setFontWeight('bold');
  Logger.log(`Prep Bay Block: wrote date header "${dateHeader}" to AI1.`);

  // Clear block (AI:AM, columns 35-39)
  const maxRows = forecastSheet.getMaxRows();
  forecastSheet.getRange(2, 35, maxRows - 1, 5).clearContent().clearFormat();

  // Write headers to row 1 (AJ-AM, columns 36-39)
  forecastSheet.getRange(1, 36, 1, 4).setValues([['Job Name', 'Order #', 'Camera Info', 'Notes']]).setFontWeight('bold');

  // Combine all jobs: today, then tomorrow, then day after tomorrow
  const allJobs = [];

  // Write today's jobs with date label in column AI
  if (todayJobs.length > 0) {
    todayJobs.forEach(job => {
      allJobs.push([dateHeader, ...job]); // Date in AI, then job data in AJ-AM
    });
  }

  // Write tomorrow's jobs with date label in column AI
  const tomorrowDateHeader = `${dayPrefixes[tomorrow.getDay()]} ${tomorrow.getMonth() + 1}/${tomorrow.getDate()}`;
  if (tomorrowJobs.length > 0) {
    tomorrowJobs.forEach(job => {
      allJobs.push([tomorrowDateHeader, ...job]); // Date in AI, then job data in AJ-AM
    });
  }

  // Write day after tomorrow's jobs with date label in column AI
  const dayAfterTomorrowDateHeader = `${dayPrefixes[dayAfterTomorrow.getDay()]} ${dayAfterTomorrow.getMonth() + 1}/${dayAfterTomorrow.getDate()}`;
  if (dayAfterTomorrowJobs.length > 0) {
    dayAfterTomorrowJobs.forEach(job => {
      allJobs.push([dayAfterTomorrowDateHeader, ...job]); // Date in AI, then job data in AJ-AM
    });
  }

  // Write all jobs to sheet
  if (allJobs.length > 0) {
    forecastSheet.getRange(2, 35, allJobs.length, 5).setValues(allJobs);
  }

  // Log summary
  Logger.log(`Prep Bay Block: wrote ${todayJobs.length} today jobs, ${tomorrowJobs.length} tomorrow jobs, ${dayAfterTomorrowJobs.length} day after tomorrow jobs to AI:AM.`);
} 