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
    const orderMatch = value.match(/\b\d{6}\b/); // Match a 6-digit order number
    const nameMatch = value.match(/"([^"]+)"/); // Match text within quotes
    const orderNumber = orderMatch ? orderMatch[0] : '';
    const jobName = nameMatch ? nameMatch[1] : '';
    return { orderNumber, jobName };
  }

  // Function to extract order number and job name from prep bay assignments
  function extractOrderAndJobName(jobString) {
    const orderMatch = jobString.match(/\b(\d+)\b/);
    const nameMatch = jobString.match(/"([^"]+)"/);
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

  // Sort sheets by date in descending order
  sortedSheets.sort((a, b) => b.sheetDate - a.sheetDate);

  // Select sheets from the most recent to find today's date and the next 7 days
  const recentSheets = sortedSheets.filter(item => item.sheetDate >= today && item.sheetDate <= sevenDaysFromNow);

  // Extract jobName and orderNumber from prep bay assignments
  Logger.log('Processing prep bay assignments from recent sheets...');
  let prepBayCount = 0;
  for (const { sheet, sheetDate } of recentSheets) {
    const sheetName = sheet.getName();
    const data = sheet.getRange('B:J').getValues(); // Get job names in column B, order numbers in column C, and notes in column J
    for (let i = 0; i < data.length; i++) {
      const jobName = data[i][0];
      const orderNumber = data[i][1];
      const note = data[i][8]; // Column J is index 8
      let jobDate = '';
      
      // Extract date from sheet name (format like 'Mon 6/2')
      const sheetDateMatch = sheetName.match(/\w+ (\d+)\/(\d+)/);
      if (sheetDateMatch) {
        const month = parseInt(sheetDateMatch[1], 10) - 1; // JavaScript months are 0-based
        const day = parseInt(sheetDateMatch[2], 10);
        const sheetDate = new Date(today.getFullYear(), month, day);
        
        // Adjust year if the sheet date is in the past
        if (sheetDate < today) {
          sheetDate.setFullYear(today.getFullYear() + 1);
        }
        jobDate = formatDate(sheetDate);
      }

      // Check if the note contains 'prep' and extract the earliest date
      if (note && note.toLowerCase().includes('prep')) {
        const dateMatches = note.match(/\d{1,2}\/\d{1,2}/g); // Match dates like 6/5 or 06/05
        if (dateMatches) {
          const earliestDate = dateMatches.reduce((earliest, current) => {
            const [currentMonth, currentDay] = current.split('/').map(Number);
            const currentDate = new Date(today.getFullYear(), currentMonth - 1, currentDay);
            return currentDate < earliest ? currentDate : earliest;
          }, new Date(today.getFullYear() + 1, 0, 1)); // Start with a future date for comparison
          jobDate = formatDate(earliestDate);
        }
      }

      if (jobName && orderNumber) {
        jobDateMap[jobName] = { date: jobDate, orderNumber: orderNumber };
        prepBayCount++;
      }
    }
  }
  Logger.log(`Completed prep bay processing: ${prepBayCount} assignments found`);

  // Get the spreadsheet and Camera sheet
  const spreadsheet = SpreadsheetApp.openById('1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw');
  const cameraSheet = spreadsheet.getSheetByName('Camera');

  // Find all rows containing "LOS ANGELES" in column A
  const data = cameraSheet.getDataRange().getValues();
  const foundLACameras = [];
  const cameraTypeCount = {}; // Object to store count of each camera type
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'LOS ANGELES') {
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
    const cameraType = data[typeRow][4]; // Column E is index 4
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

  // List of valid camera types
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
    'Indieassist 416 HD IVS '
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
      // Updated regex to handle space between BC# and barcode number
      const match = barcodeCell.match(/BC#\s*(\d+)/);
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
        if (note && note.toLowerCase().includes('crew prep')) {
          const dateMatches = note.match(/\d{1,2}\/\d{1,2}/g); // Match dates like 6/5 or 06/05
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

      // Recalculate "Today +" values and re-sort after date updates
      Logger.log('Recalculating "Today +" values after date updates...');
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
            // Calculate days difference from today with 1-day offset
            const timeDiff = parsedDate.getTime() - today.getTime();
            const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));
            
            // Apply 1-day offset: subtract 1 from the calculated difference
            const adjustedDaysDiff = daysDiff - 1;
            
            // Update the "Today +" value
            if (adjustedDaysDiff >= 0) {
              if (adjustedDaysDiff === 0) {
                row[2] = 'Today';
              } else {
                row[2] = `Today +${adjustedDaysDiff}`;
              }
              recalculatedCount++;
            }
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
        'SONY VENICE 1': 'Venice 1 Body Status'
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
                  // Set green background for RTR or Pulled status
                  if (status === 'RTR' || status === 'Pulled') {
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
    } else {
      Logger.log("Sheet 'LA Camera Forecast' not found in the active spreadsheet.");
    }
  } else {
    Logger.log('No results to write to LA Camera Forecast sheet.');
  }


} 