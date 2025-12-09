/**
 * Test Prep Bay Refresh
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

// Spreadsheet IDs (prefixed with TEST_PREP_BAY_ to avoid conflicts with other scripts)
const TEST_PREP_BAY_ASSIGNMENT_SPREADSHEET_ID = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';
const TEST_PREP_BAY_EQUIPMENT_CHART_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';
const TEST_PREP_BAY_DESTINATION_SPREADSHEET_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';
const TEST_PREP_BAY_DESTINATION_SHEET_NAME = 'Todays Prep Bays';

// Day name abbreviations
const DAY_PREFIXES = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];

/**
 * Main function to refresh prep bay data
 */
function testPrepBayRefresh() {
  Logger.log("🚀 Starting Test Prep Bay Refresh");
  
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
    
    Logger.log("✅ Test Prep Bay Refresh completed");
    
  } catch (error) {
    Logger.log(`❌ Error in testPrepBayRefresh: ${error.toString()}`);
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
    const spreadsheet = SpreadsheetApp.openById(TEST_PREP_BAY_ASSIGNMENT_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`⚠️ Sheet "${sheetName}" not found in Prep Bay Assignment`);
      return [];
    }
    
    // Read all data (columns A through I)
    // Schema: BAY (A), JOB NAME (B), ORDER (C), AGENT (D), CAMERAS (E), 1st AC (F), DP (G), PREP TECH (H), NOTES (I)
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
 * Reads Equipment Scheduling Chart data
 * @returns {Object} Object containing camera data indexed by order number
 */
function readEquipmentSchedulingData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(TEST_PREP_BAY_EQUIPMENT_CHART_ID);
    const cameraSheet = spreadsheet.getSheetByName('Camera');
    
    if (!cameraSheet) {
      Logger.log(`⚠️ Camera sheet not found in Equipment Scheduling Chart`);
      return {};
    }
    
    const data = cameraSheet.getDataRange().getValues();
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
      }
    }
    
    if (todayColumnIndex === -1) {
      Logger.log(`⚠️ Today's date column not found in Equipment Scheduling Chart`);
      return {};
    }
    
    // Map to store cameras by order number
    // Key: order number (normalized), Value: Array of {equipmentType, barcode}
    const camerasByOrder = {};
    
    // Process each row to find cameras assigned to orders
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Extract barcode from column E (index 4)
      const barcodeCell = row[4];
      let barcode = '';
      if (typeof barcodeCell === 'string') {
        const match = barcodeCell.match(/BC#\s*([A-Z0-9-]+)/);
        if (match) {
          barcode = match[1];
        }
      }
      
      if (!barcode) continue; // Skip rows without barcodes
      
      // Find camera type by looking for the first empty cell above
      let typeRow = i - 1;
      while (typeRow >= 0 && data[typeRow][0] !== '') {
        typeRow--;
      }
      const equipmentType = typeRow >= 0 ? (data[typeRow][4] || '') : ''; // Column E
      
      if (!equipmentType) continue;
      
      // Check columns F and beyond (index 5+) for order numbers
      // Look in today's column and nearby columns (today + next 7 days)
      for (let colIdx = Math.max(5, todayColumnIndex); colIdx < Math.min(row.length, todayColumnIndex + 8); colIdx++) {
        const cellValue = row[colIdx];
        if (cellValue && typeof cellValue === 'string') {
          // Extract 6-digit order numbers
          const orderMatches = cellValue.match(/\b\d{6}\b/g);
          if (orderMatches) {
            orderMatches.forEach(orderNum => {
              const normalizedOrder = orderNum.replace(/[^0-9]/g, '');
              if (!camerasByOrder[normalizedOrder]) {
                camerasByOrder[normalizedOrder] = [];
              }
              // Add camera if not already in list (avoid duplicates)
              const exists = camerasByOrder[normalizedOrder].some(cam => 
                cam.barcode === barcode && cam.equipmentType === equipmentType
              );
              if (!exists) {
                camerasByOrder[normalizedOrder].push({
                  equipmentType: equipmentType.toString().trim(),
                  barcode: barcode
                });
              }
            });
          }
        }
      }
    }
    
    Logger.log(`📚 Found cameras for ${Object.keys(camerasByOrder).length} order numbers`);
    return camerasByOrder;
    
  } catch (error) {
    Logger.log(`❌ Error reading Equipment Scheduling data: ${error.toString()}`);
    return {};
  }
}

/**
 * Writes prep bay data to the destination sheet
 * @param {Array<Object>} prepBayData - Array of prep bay assignments
 * @param {Object} equipmentData - Camera data indexed by order number
 */
function writePrepBayDataToSheet(prepBayData, equipmentData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(TEST_PREP_BAY_DESTINATION_SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(TEST_PREP_BAY_DESTINATION_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(DESTINATION_SHEET_NAME);
      Logger.log(`✅ Created new sheet: ${TEST_PREP_BAY_DESTINATION_SHEET_NAME}`);
    }
    
    // Clear the entire sheet first
    sheet.clear();
    
    // Calculate column and row positions for each prep bay
    // Prep Bay layout:
    // - Each prep bay is 4 columns wide
    // - Prep Bay 1: A1:D12 (columns 1-4)
    // - Prep Bay 2: F1:I12 (columns 6-9)
    // - Prep Bay 3: K1:N12 (columns 11-14)
    // - Pattern: Start column = (bayNumber - 1) * 5 + 1 (accounting for spacer columns)
    
    // Process each prep bay (1-22)
    for (let bayNum = 1; bayNum <= 22; bayNum++) {
      const assignment = prepBayData.find(a => a.bayNumber === bayNum);
      
      // Calculate column position
      // Within each group of 3, columns repeat: Bay 1/4/7/10/13/16/19 = col 1, Bay 2/5/8/11/14/17/20 = col 6, Bay 3/6/9/12/15/18/21 = col 11
      // Bay 22 (Kitchen) = col 1
      const positionInGroup = (bayNum - 1) % 3;
      const startCol = positionInGroup * 5 + 1; // 0->1, 1->6, 2->11
      const endCol = startCol + 3; // 4 columns wide
      
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
      
      // Write prep bay header
      const headerName = getBayDisplayName(bayNum);
      sheet.getRange(startRow, startCol + 1).setValue(headerName); // B1 equivalent
      
      if (assignment) {
        // Write job name
        sheet.getRange(startRow + 1, startCol).setValue('Job name:');
        sheet.getRange(startRow + 1, startCol + 1).setValue(assignment.jobName);
        
        // Write order number
        sheet.getRange(startRow + 2, startCol).setValue('Order #');
        sheet.getRange(startRow + 2, startCol + 1).setValue(assignment.orderNumber);
        
        // Write prep tech
        sheet.getRange(startRow + 3, startCol).setValue('Prep Tech:');
        sheet.getRange(startRow + 3, startCol + 1).setValue(assignment.prepTech);
        
        // Write cameras header
        sheet.getRange(startRow + 4, startCol).setValue('Cameras');
        
        // Get cameras for this order number
        const normalizedOrder = assignment.orderNumber.replace(/[^0-9]/g, '');
        const cameras = equipmentData[normalizedOrder] || [];
        
        // Write camera equipment types and barcodes (up to 8 cameras, rows 5-12)
        for (let i = 0; i < Math.min(cameras.length, 8); i++) {
          const camera = cameras[i];
          sheet.getRange(startRow + 4 + i, startCol + 1).setValue(camera.equipmentType); // Column B
          sheet.getRange(startRow + 4 + i, startCol + 2).setValue(camera.barcode); // Column C
        }
        
        Logger.log(`✅ Wrote data for ${headerName}: ${assignment.jobName} (Order: ${assignment.orderNumber}, ${cameras.length} cameras)`);
      } else {
        // No assignment for this bay - leave it blank but keep header
        Logger.log(`ℹ️ No assignment for ${headerName}`);
      }
    }
    
    Logger.log(`💾 Wrote prep bay data to ${TEST_PREP_BAY_DESTINATION_SHEET_NAME} sheet`);
    
  } catch (error) {
    Logger.log(`❌ Error writing to sheet: ${error.toString()}`);
    throw error;
  }
}

