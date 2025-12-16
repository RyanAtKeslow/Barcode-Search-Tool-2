/**
 * Prep Bay Refresh
 * 
 * This script reads Prep Bay Assignment data for today and populates
 * the "Todays Prep Bays" sheet with job information and camera assignments.
 * 
 * Data Sources:
 * - Prep Bay Assignment: 1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ
 * - Equipment Scheduling Chart: 1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw
 * 
 * Output:
 * - "Todays Prep Bays" sheet in: 1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw
 */

// Spreadsheet IDs
const PREP_BAY_ASSIGNMENT_SPREADSHEET_ID = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';
const PREP_BAY_EQUIPMENT_CHART_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';
const PREP_BAY_DESTINATION_SPREADSHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
const PREP_BAY_DESTINATION_SHEET_NAME = 'Todays Prep Bays';

// Day name abbreviations
const DAY_PREFIXES = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];

/**
 * Main function to refresh prep bay data
 */
function prepBayRefresh() {
  Logger.log("🚀 Starting Prep Bay Refresh");
  
  try {
    // Get today's date and sheet name
    const today = new Date();
    const todaySheetName = getTodaySheetName(today);
    Logger.log(`📅 Today's sheet name: ${todaySheetName}`);
    
    // Read Prep Bay Assignment data for today
    const prepBayData = readPrepBayDataForToday(todaySheetName);
    Logger.log(`📊 Found ${prepBayData.length} prep bay assignments`);
    
    // Read Equipment Scheduling Chart data
    const equipmentData = readEquipmentSchedulingData();
    Logger.log(`📚 Loaded Equipment Scheduling Chart data`);
    
    // Write data to destination sheet
    writePrepBayDataToSheet(prepBayData, equipmentData);
    
    Logger.log("✅ Prep Bay Refresh completed");
    
  } catch (error) {
    Logger.log(`❌ Error in prepBayRefresh: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Gets today's sheet name in format "Tues 12/9"
 * @param {Date} date - The date to get sheet name for
 * @returns {string} Sheet name like "Tues 12/9"
 */
function getTodaySheetName(date) {
  const dayPrefix = DAY_PREFIXES[date.getDay()];
  const month = date.getMonth() + 1; // 1-based
  const day = date.getDate();
  return `${dayPrefix} ${month}/${day}`;
}

/**
 * Reads Prep Bay Assignment data for today's sheet
 * @param {string} sheetName - Name of today's sheet (e.g., "Tues 12/9")
 * @returns {Array<Object>} Array of prep bay assignment objects
 */
function readPrepBayDataForToday(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(PREP_BAY_ASSIGNMENT_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`⚠️ Sheet "${sheetName}" not found in Prep Bay Assignment`);
      return [];
    }
    
    // Read all data from Prep Bay Assignment workbook (NO column shifts in this source workbook)
    // Schema: BAY (A), JOB NAME (B), ORDER (C), AGENT (D), CAMERAS (E), 1st AC (F), DP (G), PREP TECH (H), NOTES (I)
    // Prep Tech is read from column H (index 7) - this workbook has NOT been modified with column shifts
    const data = sheet.getDataRange().getValues();
    
    const prepBayAssignments = [];
    
    // Process each row (skip header row if present)
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const bay = row[0] ? row[0].toString().trim() : ''; // Column A
      const jobName = row[1] ? row[1].toString().trim() : ''; // Column B
      const orderNumber = row[2] ? row[2].toString().trim() : ''; // Column C
      const prepTech = row[7] ? row[7].toString().trim() : ''; // Column H (Prep Tech)
      
      // Skip empty rows or header rows
      if (!bay || bay.toUpperCase() === 'BAY' || !jobName) {
        continue;
      }
      
      // Normalize bay number/name
      const bayNumber = normalizeBayNumber(bay);
      if (bayNumber === null) {
        continue; // Skip invalid bay numbers
      }
      
      prepBayAssignments.push({
        bayNumber: bayNumber,
        bayName: bay,
        jobName: jobName,
        orderNumber: orderNumber,
        prepTech: prepTech
      });
    }
    
    // Sort by bay number
    prepBayAssignments.sort((a, b) => a.bayNumber - b.bayNumber);
    
    return prepBayAssignments;
    
  } catch (error) {
    Logger.log(`❌ Error reading Prep Bay data: ${error.toString()}`);
    return [];
  }
}

/**
 * Normalizes bay number/name to numeric value (1-22)
 * @param {string} bay - Bay identifier (e.g., "1", "BL 1", "BL 2", "KTN")
 * @returns {number|null} Bay number (1-22) or null if invalid
 */
function normalizeBayNumber(bay) {
  const bayStr = bay.toString().trim().toUpperCase();
  
  // Check for numbered bays (1-19)
  const numberMatch = bayStr.match(/^(\d+)$/);
  if (numberMatch) {
    const num = parseInt(numberMatch[1], 10);
    if (num >= 1 && num <= 19) {
      return num;
    }
  }
  
  // Check for Backlot 1 (BL 1)
  if (bayStr === 'BL 1' || bayStr === 'BACKLOT 1') {
    return 20;
  }
  
  // Check for Backlot 2 (BL 2)
  if (bayStr === 'BL 2' || bayStr === 'BACKLOT 2') {
    return 21;
  }
  
  // Check for Kitchen (KTN)
  if (bayStr === 'KTN' || bayStr === 'KITCHEN') {
    return 22;
  }
  
  return null;
}

/**
 * Gets display name for a bay number
 * @param {number} bayNumber - Bay number (1-22)
 * @returns {string} Display name like "PREP BAY 1" or "BACKLOT 1"
 */
function getBayDisplayName(bayNumber) {
  if (bayNumber >= 1 && bayNumber <= 19) {
    return `PREP BAY ${bayNumber}`;
  } else if (bayNumber === 20) {
    return 'BACKLOT 1';
  } else if (bayNumber === 21) {
    return 'BACKLOT 2';
  } else if (bayNumber === 22) {
    return 'KITCHEN';
  }
  return `PREP BAY ${bayNumber}`;
}

/**
 * Reads Equipment Scheduling Chart data using Camera Forecast logic
 * Finds cameras assigned to orders by checking LOS ANGELES rows and today's date column
 * Reads from both "Camera" and "Consignor Use Only" sheets
 * @returns {Object} Object containing camera data indexed by order number
 */
function readEquipmentSchedulingData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PREP_BAY_EQUIPMENT_CHART_ID);
    
    // Map to store cameras by order number (merged from both sheets)
    // Key: order number (normalized), Value: Array of {equipmentType, barcode}
    const camerasByOrder = {};
    
    // Process Camera sheet
    const cameraSheet = spreadsheet.getSheetByName('Camera');
    if (cameraSheet) {
      Logger.log(`📹 Processing Camera sheet...`);
      const cameraData = processEquipmentSheet(cameraSheet, 'Camera');
      // Merge camera data into main object
      for (const [orderNumber, cameras] of Object.entries(cameraData)) {
        if (!camerasByOrder[orderNumber]) {
          camerasByOrder[orderNumber] = [];
        }
        // Add cameras, avoiding duplicates by barcode
        for (const camera of cameras) {
          const exists = camerasByOrder[orderNumber].some(cam => cam.barcode === camera.barcode);
          if (!exists) {
            camerasByOrder[orderNumber].push(camera);
          }
        }
      }
    } else {
      Logger.log(`⚠️ Camera sheet not found in Equipment Scheduling Chart`);
    }
    
    // Process Consignor Use Only sheet
    const consignorSheet = spreadsheet.getSheetByName('Consignor Use Only');
    if (consignorSheet) {
      Logger.log(`📹 Processing Consignor Use Only sheet...`);
      const consignorData = processEquipmentSheet(consignorSheet, 'Consignor Use Only');
      // Merge consignor data into main object
      for (const [orderNumber, cameras] of Object.entries(consignorData)) {
        if (!camerasByOrder[orderNumber]) {
          camerasByOrder[orderNumber] = [];
        }
        // Add cameras, avoiding duplicates by barcode
        for (const camera of cameras) {
          const exists = camerasByOrder[orderNumber].some(cam => cam.barcode === camera.barcode);
          if (!exists) {
            camerasByOrder[orderNumber].push(camera);
          }
        }
      }
    } else {
      Logger.log(`⚠️ Consignor Use Only sheet not found in Equipment Scheduling Chart`);
    }
    
    Logger.log(`📚 Found cameras for ${Object.keys(camerasByOrder).length} order numbers`);
    
    // Log camera count for each order
    for (const [orderNumber, cameras] of Object.entries(camerasByOrder)) {
      Logger.log(`  Order ${orderNumber}: ${cameras.length} camera(s) - ${cameras.map(c => c.barcode).join(', ')}`);
    }
    
    return camerasByOrder;
    
  } catch (error) {
    Logger.log(`❌ Error reading Equipment Scheduling data: ${error.toString()}`);
    return {};
  }
}

/**
 * Processes a single equipment sheet (Camera or Consignor Use Only)
 * Finds cameras assigned to orders by checking LOS ANGELES rows and today's date column
 * @param {Sheet} sheet - The sheet to process
 * @param {string} sheetName - Name of the sheet (for logging)
 * @returns {Object} Object containing camera data indexed by order number
 */
function processEquipmentSheet(sheet, sheetName) {
  try {
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // Find today's date column
    const today = new Date();
    let todayColumnIndex = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      const cell = headerRow[i];
      if (cell instanceof Date) {
        if (
          cell.getFullYear() === today.getFullYear() &&
          cell.getMonth() === today.getMonth() &&
          cell.getDate() === today.getDate()
        ) {
          todayColumnIndex = i;
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
            todayColumnIndex = i;
            break;
          }
        }
      }
    }
    
    if (todayColumnIndex === -1) {
      Logger.log(`⚠️ Today's date column not found in ${sheetName} sheet`);
      return {};
    }
    
    Logger.log(`📅 Found today's date in column ${todayColumnIndex + 1} of ${sheetName} sheet`);
    
    // Get backgrounds for all rows to check for valid booking colors
    const validTodayCellBackgrounds = [
      '#ffffff', // white
      '#f9ff71', // yellow
      '#66ff75', // green
      '#4a86e8', // blue
      '#ff7171', // red
      '#00ffff'  // cyan
    ];
    
    // Find all rows containing "LOS ANGELES" in column A
    const foundLACameras = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'LOS ANGELES') {
        foundLACameras.push(i + 1); // 1-based row number
      }
    }
    
    Logger.log(`📹 Found ${foundLACameras.length} LOS ANGELES camera rows in ${sheetName} sheet`);
    
    // Get backgrounds for LA camera rows
    const backgrounds = sheet.getRange(1, todayColumnIndex + 1, data.length, 1).getBackgrounds();
    
    // Map to store cameras by order number
    // Key: order number (normalized), Value: Array of {equipmentType, barcode}
    const camerasByOrder = {};
    
    // Process each LA camera row
    for (const rowNum of foundLACameras) {
      const rowIdx = rowNum - 1; // Convert to 0-based
      const row = data[rowIdx];
      
      // Check if today's cell has a valid background color (indicating booking)
      const cellBg = backgrounds[rowIdx][0];
      if (!validTodayCellBackgrounds.includes(cellBg)) {
        continue; // Skip if not a valid booking color
      }
      
      // Extract barcode from column E (index 4) first - we need this for all cameras
      const barcodeCell = row[4];
      let barcode = '';
      if (typeof barcodeCell === 'string') {
        const match = barcodeCell.match(/BC#\s*([A-Z0-9-]+)/);
        if (match) {
          barcode = match[1];
        }
      }
      
      if (!barcode) {
        continue; // Skip rows without barcodes
      }
      
      // Find camera type by looking for the first empty cell above
      let typeRow = rowIdx - 1;
      while (typeRow >= 0 && data[typeRow][0] !== '') {
        typeRow--;
      }
      const equipmentType = typeRow >= 0 ? (data[typeRow][4] || '') : ''; // Column E
      
      if (!equipmentType) {
        continue; // Skip if no camera type found
      }
      
      // Search for order numbers in today's column
      // Special case: If today's cell has red background (#ff7171) and is blank, search left for order number
      const orderNumbersFound = new Set();
      const todayCellValue = row[todayColumnIndex];
      const isRedBackground = cellBg === '#ff7171';
      const isTodayCellBlank = !todayCellValue || (typeof todayCellValue === 'string' && todayCellValue.trim() === '');
      
      if (isRedBackground && isTodayCellBlank) {
        // Red background with blank cell: search leftward for order number
        Logger.log(`🔍 Red background with blank cell for barcode ${barcode} in ${sheetName}, searching left for order number...`);
        for (let colIdx = todayColumnIndex - 1; colIdx >= 6; colIdx--) { // Start from column before today, go back to column G (index 6)
          const leftCellValue = row[colIdx];
          if (leftCellValue && typeof leftCellValue === 'string' && leftCellValue.trim() !== '') {
            // Check if this cell has an order number
            const orderMatches = leftCellValue.match(/\b(\d{6})\b/g);
            if (orderMatches && orderMatches.length > 0) {
              // Found order number(s) - use these
              orderMatches.forEach(ord => orderNumbersFound.add(ord.replace(/[^0-9]/g, '')));
              Logger.log(`  ✅ Found order number(s) ${Array.from(orderNumbersFound).join(', ')} in column ${colIdx + 1}`);
              break; // Stop searching once we find an order number
            }
          }
        }
      } else if (todayCellValue && typeof todayCellValue === 'string' && todayCellValue.trim() !== '') {
        // Today's cell has content - extract order numbers from it
        const todayOrderMatches = todayCellValue.match(/\b(\d{6})\b/g);
        if (todayOrderMatches) {
          todayOrderMatches.forEach(ord => orderNumbersFound.add(ord.replace(/[^0-9]/g, '')));
        }
      }
      
      // Add this camera to ALL order numbers found
      if (orderNumbersFound.size > 0) {
        for (const normalizedOrder of orderNumbersFound) {
          // Add camera to the order's list
          if (!camerasByOrder[normalizedOrder]) {
            camerasByOrder[normalizedOrder] = [];
          }
          
          // Add camera if not already in list (avoid duplicates by barcode only)
          const exists = camerasByOrder[normalizedOrder].some(cam => 
            cam.barcode === barcode
          );
          if (!exists) {
            camerasByOrder[normalizedOrder].push({
              equipmentType: equipmentType.toString().trim(),
              barcode: barcode
            });
            Logger.log(`📹 Added camera from ${sheetName}: Order ${normalizedOrder}, Type: ${equipmentType}, Barcode: ${barcode}`);
          }
        }
      }
    }
    
    return camerasByOrder;
    
  } catch (error) {
    Logger.log(`❌ Error processing ${sheetName} sheet: ${error.toString()}`);
    return {};
  }
}

/**
 * Writes prep bay data to the destination sheet ("Todays Prep Bays" in Digital Camera Service Forms workbook)
 * Shows all cameras for an order in all prep bays with that order number
 * 
 * NOTE: Column shifts (new "Pulled?" columns D, J, P) were ONLY done in this destination sheet.
 * The source Prep Bay Assignment workbook has NO column shifts.
 * 
 * @param {Array<Object>} prepBayData - Array of prep bay assignments
 * @param {Object} equipmentData - Camera data indexed by order number
 */
function writePrepBayDataToSheet(prepBayData, equipmentData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(PREP_BAY_DESTINATION_SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(PREP_BAY_DESTINATION_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(PREP_BAY_DESTINATION_SHEET_NAME);
      Logger.log(`✅ Created new sheet: ${PREP_BAY_DESTINATION_SHEET_NAME}`);
    }
    
    // Process each prep bay (1-22)
    for (let bayNum = 1; bayNum <= 22; bayNum++) {
      const assignment = prepBayData.find(a => a.bayNumber === bayNum);
      
      // Calculate column position
      // Within each group of 3, columns repeat: Bay 1/4/7/10/13/16/19 = col 1 (A), Bay 2/5/8/11/14/17/20 = col 7 (G), Bay 3/6/9/12/15/18/21 = col 13 (M)
      // Bay 22 (Kitchen) = col 1 (A)
      // Note: Column F (6) is blank between Prep Bay 1 and Prep Bay 2
      const positionInGroup = (bayNum - 1) % 3;
      const startCol = positionInGroup * 6 + 1; // 0->1 (A), 1->7 (G), 2->13 (M)
      
      // Calculate row position
      // Each group of 3 prep bays shares the same starting row
      // Group 0 (Bays 1-3): rows 1-12
      // Group 1 (Bays 4-6): rows 14-25 (row 13 is blank)
      // Group 2 (Bays 7-9): rows 27-38 (row 26 is blank)
      // Group 3 (Bays 10-12): rows 40-51 (row 39 is blank)
      // Group 4 (Bays 13-15): rows 53-64 (row 52 is blank)
      // Group 5 (Bays 16-18): rows 66-77 (row 65 is blank)
      // Group 6 (Bays 19-21): rows 79-90 (row 78 is blank)
      // Group 7 (Bay 22): rows 92-103 (row 91 is blank)
      
      // Calculate which group this bay belongs to (0-based)
      const group = Math.floor((bayNum - 1) / 3);
      // Each group starts 13 rows after the previous (12 rows + 1 blank row)
      const startRow = 1 + group * 13;
      
      // Layout for each prep bay (using Prep Bay 1 as example):
      // - A1: Datum (empty, untouched)
      // - B1: "Prep Bay X" header (untouched)
      // - A1:E1: Untouched (nothing printed here)
      // - A2:A5: Left-side headers (untouched)
      // - B2: Job Name (script writes here)
      // - B3: Order Number (script writes here)
      // - B4: Prep Tech name (script writes here, from Prep Bay Assignment Sheet column H)
      // - B5:B12: Camera names (script writes here)
      // - C5:C12: Barcodes (script writes here)
      // - C1:E4: Untouched
      // - D5:E12: Untouched (D5:D12 are "Pulled?" checkboxes)
      // - Borders: Column F and Row 13 (nothing printed)
      
      // Clear only the cells we will write to: B2:B12 and C5:C12
      // Clear B2:B12 (Job Name, Order Number, Prep Tech, Camera names)
      sheet.getRange(startRow + 1, startCol + 1, 11, 1).clearContent();
      
      // Clear C5:C12 (Barcodes only)
      // Note: D5:D12 (checkboxes) are intentionally left untouched
      sheet.getRange(startRow + 4, startCol + 2, 8, 1).clearContent();
      
      if (assignment) {
        // Write job name to B2
        sheet.getRange(startRow + 1, startCol + 1).setValue(assignment.jobName);
        
        // Write order number to B3
        sheet.getRange(startRow + 2, startCol + 1).setValue(assignment.orderNumber);
        
        // Write prep tech to B4 (read from Prep Bay Assignment Sheet column H, index 7)
        sheet.getRange(startRow + 3, startCol + 1).setValue(assignment.prepTech);
        
        // Get all cameras scheduled for this order number TODAY
        // All cameras for the order are shown in ALL prep bays with this order number
        const normalizedOrder = assignment.orderNumber.replace(/[^0-9]/g, '');
        const cameras = equipmentData[normalizedOrder] || [];
        
        // Write each camera body to its own row in B5:B12 (up to 8 cameras)
        // Each camera body gets its own row: equipment type in column B, barcode in column C
        for (let i = 0; i < Math.min(cameras.length, 8); i++) {
          const camera = cameras[i];
          // Write camera equipment type to B5, B6, B7, etc.
          sheet.getRange(startRow + 4 + i, startCol + 1).setValue(camera.equipmentType);
          // Write corresponding barcode to C5, C6, C7, etc.
          sheet.getRange(startRow + 4 + i, startCol + 2).setValue(camera.barcode);
        }
        
        // Clear remaining cells in B5:B12 and C5:C12 if there are fewer than 8 cameras
        if (cameras.length < 8) {
          // Clear unused cells (beyond the number of cameras)
          for (let i = cameras.length; i < 8; i++) {
            sheet.getRange(startRow + 4 + i, startCol + 1).setValue(''); // Clear unused B cells
            sheet.getRange(startRow + 4 + i, startCol + 2).setValue(''); // Clear unused C cells
          }
        }
        
        const headerName = getBayDisplayName(bayNum);
        Logger.log(`✅ Wrote data for ${headerName}: ${assignment.jobName} (Order: ${assignment.orderNumber}, ${cameras.length} camera(s) scheduled for today)`);
      } else {
        // No assignment for this bay - cells already cleared above
        const headerName = getBayDisplayName(bayNum);
        Logger.log(`ℹ️ No assignment for ${headerName} - cleared data cells`);
      }
    }
    
    Logger.log(`💾 Wrote prep bay data to ${PREP_BAY_DESTINATION_SHEET_NAME} sheet`);
    
  } catch (error) {
    Logger.log(`❌ Error writing to sheet: ${error.toString()}`);
    throw error;
  }
}

/**
 * Clears all prep bay data cells for all 22 prep bays
 * For each prep bay, clears B2:B12 (Job Name, Order Number, Prep Tech, Camera names) and C5:C12 (Barcodes)
 * Does NOT clear: C1:E4, D5:E12 (D5:D12 are "Pulled?" checkboxes), or any headers
 */
function clearAllPrepBays() {
  try {
    Logger.log("🧹 Starting to clear all prep bays");
    
    const spreadsheet = SpreadsheetApp.openById(PREP_BAY_DESTINATION_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(PREP_BAY_DESTINATION_SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`⚠️ Sheet "${PREP_BAY_DESTINATION_SHEET_NAME}" not found`);
      return;
    }
    
    // Process each prep bay (1-22)
    for (let bayNum = 1; bayNum <= 22; bayNum++) {
      // Calculate column position
      // Within each group of 3, columns repeat: Bay 1/4/7/10/13/16/19 = col 1 (A), Bay 2/5/8/11/14/17/20 = col 7 (G), Bay 3/6/9/12/15/18/21 = col 13 (M)
      // Bay 22 (Kitchen) = col 1 (A)
      // Note: Column F (6) is blank between Prep Bay 1 and Prep Bay 2
      const positionInGroup = (bayNum - 1) % 3;
      const startCol = positionInGroup * 6 + 1; // 0->1 (A), 1->7 (G), 2->13 (M)
      
      // Calculate row position
      // Each group of 3 prep bays shares the same starting row
      // Group 0 (Bays 1-3): rows 1-12
      // Group 1 (Bays 4-6): rows 14-25 (row 13 is blank)
      // Group 2 (Bays 7-9): rows 27-38 (row 26 is blank)
      // Group 3 (Bays 10-12): rows 40-51 (row 39 is blank)
      // Group 4 (Bays 13-15): rows 53-64 (row 52 is blank)
      // Group 5 (Bays 16-18): rows 66-77 (row 65 is blank)
      // Group 6 (Bays 19-21): rows 79-90 (row 78 is blank)
      // Group 7 (Bay 22): rows 92-103 (row 91 is blank)
      
      // Calculate which group this bay belongs to (0-based)
      const group = Math.floor((bayNum - 1) / 3);
      // Each group starts 13 rows after the previous (12 rows + 1 blank row)
      const startRow = 1 + group * 13;
      
      // Layout for each prep bay (same as writePrepBayDataToSheet):
      // - B2: Job Name (clear)
      // - B3: Order Number (clear)
      // - B4: Prep Tech (clear)
      // - B5:B12: Camera names (clear)
      // - C5:C12: Barcodes (clear)
      // - D5:D12: "Pulled?" checkboxes (DO NOT CLEAR - these are in columns D, J, P for different prep bays)
      
      // Clear B2:B12 (Job Name, Order Number, Prep Tech, Camera names)
      // Skip if column B would fall on a "Pulled?" checkbox column (D=4, J=10, P=16)
      const columnB = startCol + 1;
      if (columnB !== 4 && columnB !== 10 && columnB !== 16) {
        sheet.getRange(startRow + 1, startCol + 1, 11, 1).clearContent();
      }
      
      // Clear C5:C12 (Barcodes only)
      // Skip if column C would fall on a "Pulled?" checkbox column (D=4, J=10, P=16)
      const columnC = startCol + 2;
      if (columnC !== 4 && columnC !== 10 && columnC !== 16) {
        sheet.getRange(startRow + 4, startCol + 2, 8, 1).clearContent();
      }
      
      // Note: "Pulled?" checkbox columns are:
      // - Prep Bay 1: Column D (4) = startCol (1) + 3
      // - Prep Bay 2: Column J (10) = startCol (7) + 3
      // - Prep Bay 3: Column P (16) = startCol (13) + 3
      // These are intentionally skipped and never cleared
      
      const headerName = getBayDisplayName(bayNum);
      Logger.log(`🧹 Cleared ${headerName}`);
    }
    
    Logger.log("✅ Successfully cleared all prep bays");
    
  } catch (error) {
    Logger.log(`❌ Error clearing prep bays: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

