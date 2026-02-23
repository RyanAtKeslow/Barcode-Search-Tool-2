/**
 * LA Prep Floor & Camera Service Status — Prep Bay Schema Test
 *
 * Writes the v2 layout (job-centric, vertical) to the "Prep Bay Schema" sheet
 * for testing. Uses sample data only; no live Prep Bay or Equipment Chart reads yet.
 *
 * Workbook: LA Prep Floor & Camera Service Status
 * ID: 1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0
 * Sheet: Prep Bay Schema
 */

const LA_PREP_STATUS_WORKBOOK_ID = '1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0';
const PREP_BAY_SCHEMA_SHEET_NAME = 'Prep Bay Schema';
const PREP_BAY_ASSIGNMENT_SPREADSHEET_ID = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';
const PREP_BAY_EQUIPMENT_CHART_ID = '1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw';
/** Workbook that contains "Camera Bodies Only" sheet (same as Prep Bay Refresh destination) */
const CAMERA_BODIES_ONLY_WORKBOOK_ID = '1FYA76P4B7vFUCDmxDwc6Ly6-tm7F6f5c5v0eNYjgwKw';

/** Forecast sheet names in the LA Prep workbook and their day offset (0 = today) */
const PREP_FORECAST_SHEETS = [
  { name: 'Prep Today', daysOffset: 0 },
  { name: 'Prep Tomorrow', daysOffset: 1 },
  { name: 'Prep Two Days Out', daysOffset: 2 },
  { name: 'Prep Three Days Out', daysOffset: 3 },
  { name: 'Prep Four Days Out', daysOffset: 4 }
];

/** Day name abbreviations for sheet names (same as Prep Bay Refresh) */
const DAY_PREFIXES = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];

/** Equipment Scheduling Chart status colors (same as Prep Bay Refresh) */
const STATUS_COLORS = {
  ON_JOB_AT: '#fbacff', ON_JOB_CH: '#ffc566', ON_JOB_NO: '#b28701', ON_JOB_ABQ: '#e65d16',
  ON_JOB_LA: '#cb7fff', ON_JOB_TO: '#bbd6ff', ON_JOB_VN: '#b1ffe9',
  IN_REPAIR: '#ff4444', CONSIGNOR: '#00ffff', TURNAROUND_MULTIDAY: '#ff7171',
  GEAR_TRANSFER: '#4a86e8', DO_NOT_RESCHEDULE: '#bdbdbd', CONFIRMED_JOB: '#f9ff71', PENDING_JOB: '#66ff75'
};
const VALID_TODAY_BACKGROUNDS_FOR_PREP_BAY = [
  '#ffffff', STATUS_COLORS.CONFIRMED_JOB, STATUS_COLORS.TURNAROUND_MULTIDAY, STATUS_COLORS.CONSIGNOR,
  STATUS_COLORS.GEAR_TRANSFER, STATUS_COLORS.DO_NOT_RESCHEDULE, STATUS_COLORS.IN_REPAIR, STATUS_COLORS.PENDING_JOB
];

/** Row count per job block (job header 6 + blank 1 + eq header 1 + categories 10 + blank 1 + sub header 1 + sub row 1 + blank 1) */
const ROWS_PER_JOB_BLOCK = 22;

/** Default formatting (used when Settings sheet is missing or a value is blank) */
const FMT_DEFAULTS = {
  jobHeaderBg: '#e8f0fe',
  jobNameValueSize: 14,
  jobNameValueColor: '#1a73e8',
  labelColor: '#5f6368',
  labelSize: 11,
  valueSize: 11,
  tableHeaderBg: '#344a5e',
  tableHeaderFg: '#ffffff',
  tableHeaderSize: 11,
  categoryBg: '#f1f3f4',
  subHeaderBg: '#344a5e',
  subHeaderFg: '#ffffff',
  borderColor: '#dadce0',
  rowHeightJobName: 40,
  rowHeightLabel: 28,
  rowHeightTableHeader: 32,
  rowHeightCategory: 24,
  colWidthLabel: 120,
  colWidthValue: 220
};

/** Equipment categories (v2 schema) — rows can grow dynamically per job */
const EQUIPMENT_CATEGORIES = [
  'Cameras',
  'Lenses',
  'Heads',
  'Focus',
  'Matte Boxes',
  'Monitors',
  'Media',
  'Wireless Video',
  'Dir. Viewfinder',
  'Ungrouped'
];

// ---------------------------------------------------------------------------
// Prep Bay Assignment: sheet-by-date lookup and order number cell background
// ---------------------------------------------------------------------------

/**
 * Sheet name for a date in Prep Bay Assignment format (e.g. "Tues 12/9").
 * @param {Date} date
 * @returns {string}
 */
function getTodaySheetName(date) {
  const dayPrefix = DAY_PREFIXES[date.getDay()];
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return dayPrefix + ' ' + month + '/' + day;
}

function normalizeDayAbbreviation(dayAbbr) {
  if (!dayAbbr) return '';
  const upper = String(dayAbbr).toUpperCase();
  if (upper === 'TUE' || upper === 'TUES') return 'Tues';
  if (upper === 'THUR' || upper === 'THURS') return 'Thurs';
  return dayAbbr;
}

function normalizeSheetName(name) {
  if (!name) return '';
  let normalized = String(name).replace(/[^\x00-\x7F\s\/]/g, '').trim();
  const dayAbbrMatch = normalized.match(/^(\w+)\s+(\d+\/\d+)/);
  if (dayAbbrMatch) {
    const normalizedDay = normalizeDayAbbreviation(dayAbbrMatch[1]);
    normalized = normalizedDay + ' ' + dayAbbrMatch[2];
  }
  return normalized;
}

/**
 * Finds a sheet by name pattern (handles emojis / day abbreviation variants).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} expectedName - e.g. "Mon 12/22"
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findPrepBaySheetByNamePattern(spreadsheet, expectedName) {
  const exact = spreadsheet.getSheetByName(expectedName);
  if (exact) return exact;
  const normalizedExpected = normalizeSheetName(expectedName);
  if (!normalizedExpected) return null;
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const normalized = normalizeSheetName(sheets[i].getName());
    if (normalized === normalizedExpected) return sheets[i];
  }
  return null;
}

/**
 * Gets the background color of the order number cell (column C) in the Prep Bay
 * Assignment sheet for the given date sheet and order number.
 * Schema: A=Bay, B=Job Name, C=Order, ...
 * @param {string} sheetName - Date sheet name (e.g. "Fri 2/20")
 * @param {string} orderNumber - Order number (e.g. "881951"); matched after normalizing to digits
 * @returns {string|null} Hex background color (e.g. "#e8f0fe") or null if not found
 */
function getOrderNumberBackgroundFromPrepBay(sheetName, orderNumber) {
  try {
    const ss = SpreadsheetApp.openById(PREP_BAY_ASSIGNMENT_SPREADSHEET_ID);
    const sheet = findPrepBaySheetByNamePattern(ss, sheetName);
    if (!sheet) return null;
    const orderNorm = String(orderNumber || '').replace(/[^0-9]/g, '');
    if (!orderNorm) return null;
    const data = sheet.getDataRange();
    const numRows = data.getNumRows();
    const colC = 3;
    const values = sheet.getRange(1, colC, numRows, colC).getValues();
    const backgrounds = sheet.getRange(1, colC, numRows, colC).getBackgrounds();
    for (let i = 0; i < values.length; i++) {
      const cellOrder = String(values[i][0] || '').replace(/[^0-9]/g, '');
      if (cellOrder === orderNorm) {
        const bg = backgrounds[i][0];
        return bg && String(bg).trim() ? bg : null;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getOrderNumberBackgroundFromPrepBay: ' + e.message);
    return null;
  }
}

/**
 * Normalizes barcode for comparison (trim, uppercase).
 * @param {string} barcode
 * @returns {string}
 */
function normalizeBarcode(barcode) {
  if (!barcode) return '';
  return String(barcode).trim().toUpperCase();
}

/**
 * Normalizes bay to numeric 1-22 (same as Prep Bay Refresh).
 * @param {string} bay
 * @returns {number|null}
 */
function normalizeBayNumber(bay) {
  const s = String(bay || '').trim().toUpperCase();
  const numMatch = s.match(/^(\d+)$/);
  if (numMatch) {
    const n = parseInt(numMatch[1], 10);
    if (n >= 1 && n <= 19) return n;
  }
  if (s === 'BL 1' || s === 'BACKLOT 1') return 20;
  if (s === 'BL 2' || s === 'BACKLOT 2') return 21;
  if (s === 'KTN' || s === 'KITCHEN') return 22;
  return null;
}

/**
 * Reads Prep Bay Assignment for a given date sheet name.
 * Schema: A=Bay, B=Job Name, C=Order, D=Agent, E=Cameras, F=1st AC, G=DP, H=Prep Tech, I=Notes.
 * @param {string} sheetName - e.g. "Fri 2/20"
 * @returns {Array<Object>} [{ bayNumber, bayName, jobName, orderNumber, prepTech, notes }, ...]
 */
function readPrepBayDataForDate(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(PREP_BAY_ASSIGNMENT_SPREADSHEET_ID);
    const sheet = findPrepBaySheetByNamePattern(ss, sheetName);
    if (!sheet) {
      Logger.log('Prep Bay sheet not found: ' + sheetName);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    const out = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const bay = row[0] ? String(row[0]).trim() : '';
      const jobName = row[1] ? String(row[1]).trim() : '';
      const orderNumber = row[2] ? String(row[2]).trim() : '';
      const prepTech = row[7] ? String(row[7]).trim() : '';
      const notes = row[8] ? String(row[8]).trim() : '';
      if (!bay || bay.toUpperCase() === 'BAY' || !jobName) continue;
      const bayNumber = normalizeBayNumber(bay);
      if (bayNumber == null) continue;
      out.push({ bayNumber: bayNumber, bayName: bay, jobName: jobName, orderNumber: orderNumber, prepTech: prepTech, notes: notes });
    }
    out.sort(function (a, b) { return a.bayNumber - b.bayNumber; });
    return out;
  } catch (e) {
    Logger.log('readPrepBayDataForDate: ' + e.message);
    return [];
  }
}

/**
 * Reads Camera Bodies Only lookup from the workbook that has that sheet (barcode -> name from column D).
 * @returns {Object}
 */
function readCameraBodiesOnlyLookup() {
  try {
    const ss = SpreadsheetApp.openById(CAMERA_BODIES_ONLY_WORKBOOK_ID);
    const sheet = ss.getSheetByName('Camera Bodies Only');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    const lookup = {};
    for (let i = 0; i < data.length; i++) {
      const barcode = data[i][8];
      const name = data[i][3];
      if (barcode && name) {
        const nb = normalizeBarcode(barcode);
        if (nb) lookup[nb] = String(name).trim();
      }
    }
    return lookup;
  } catch (e) {
    Logger.log('readCameraBodiesOnlyLookup: ' + e.message);
    return {};
  }
}

/**
 * Processes one equipment sheet (Camera or Consignor Use Only) for a target date; LOS ANGELES rows only.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} sheetName
 * @param {Object|null} cameraBodiesLookup
 * @param {Date} targetDate
 * @returns {Object} orderNumber -> [{ equipmentType, barcode }]
 */
function processEquipmentSheet(sheet, sheetName, cameraBodiesLookup, targetDate) {
  try {
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    let targetDateColumnIndex = -1;
    for (let i = 0; i < headerRow.length; i++) {
      const cell = headerRow[i];
      if (cell instanceof Date) {
        if (cell.getFullYear() === targetDate.getFullYear() && cell.getMonth() === targetDate.getMonth() && cell.getDate() === targetDate.getDate()) {
          targetDateColumnIndex = i;
          break;
        }
      } else if (typeof cell === 'string') {
        const parts = cell.split('/');
        if (parts.length === 3) {
          const m = parseInt(parts[0], 10), d = parseInt(parts[1], 10), y = parseInt(parts[2], 10);
          if (y === targetDate.getFullYear() && m === targetDate.getMonth() + 1 && d === targetDate.getDate()) {
            targetDateColumnIndex = i;
            break;
          }
        }
      }
    }
    if (targetDateColumnIndex === -1) return {};
    const foundLACameras = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'LOS ANGELES') foundLACameras.push(i + 1);
    }
    const backgrounds = sheet.getRange(1, targetDateColumnIndex + 1, data.length, 1).getBackgrounds();
    const camerasByOrder = {};
    const isConsignor = sheetName === 'Consignor Use Only';
    for (let ri = 0; ri < foundLACameras.length; ri++) {
      const rowIdx = foundLACameras[ri] - 1;
      const row = data[rowIdx];
      const cellBg = (backgrounds[rowIdx] && backgrounds[rowIdx][0]) ? backgrounds[rowIdx][0] : '';
      if (VALID_TODAY_BACKGROUNDS_FOR_PREP_BAY.indexOf(cellBg) === -1) continue;
      let barcode = '';
      if (typeof row[4] === 'string') {
        const match = row[4].match(/BC#\s*([A-Z0-9-]+)/);
        if (match) barcode = match[1];
      }
      if (!barcode) continue;
      let equipmentType = '';
      if (isConsignor && cameraBodiesLookup) {
        const name = cameraBodiesLookup[normalizeBarcode(barcode)];
        if (!name) continue;
        equipmentType = name;
      } else {
        let typeRow = rowIdx - 1;
        while (typeRow >= 0 && data[typeRow][0] !== '') typeRow--;
        equipmentType = typeRow >= 0 ? (data[typeRow][4] || '') : '';
        if (!equipmentType) continue;
        equipmentType = String(equipmentType).trim();
      }
      const targetDateCellValue = row[targetDateColumnIndex];
      const orderNumbersFound = new Set();
      const isTurnaround = cellBg === STATUS_COLORS.TURNAROUND_MULTIDAY;
      const isConfirmedJob = cellBg === STATUS_COLORS.CONFIRMED_JOB;
      const isConsignorColor = cellBg === STATUS_COLORS.CONSIGNOR;
      const isDoNotReschedule = cellBg === STATUS_COLORS.DO_NOT_RESCHEDULE;
      const isGearTransfer = cellBg === STATUS_COLORS.GEAR_TRANSFER;
      const isInRepair = cellBg === STATUS_COLORS.IN_REPAIR;
      const isPendingJob = cellBg === STATUS_COLORS.PENDING_JOB;
      const isBlank = !targetDateCellValue || (typeof targetDateCellValue === 'string' && targetDateCellValue.trim() === '');
      function searchLeft() {
        for (let colIdx = targetDateColumnIndex - 1; colIdx >= 6; colIdx--) {
          const v = row[colIdx];
          if (v && typeof v === 'string' && v.trim() !== '') {
            const matches = v.match(/\b(\d{6})\b/g);
            if (matches) {
              matches.forEach(function (ord) { orderNumbersFound.add(ord.replace(/[^0-9]/g, '')); });
              break;
            }
          }
        }
      }
      if (isTurnaround && isBlank) searchLeft();
      else if (isConfirmedJob) searchLeft();
      else if (isConsignorColor) searchLeft();
      else if (isDoNotReschedule) searchLeft();
      else if (isGearTransfer && isBlank) searchLeft();
      else if (isInRepair) searchLeft();
      else if (isPendingJob) searchLeft();
      else if (targetDateCellValue && typeof targetDateCellValue === 'string' && targetDateCellValue.trim() !== '') {
        const matches = targetDateCellValue.match(/\b(\d{6})\b/g);
        if (matches) matches.forEach(function (ord) { orderNumbersFound.add(ord.replace(/[^0-9]/g, '')); });
      }
      let displayType = equipmentType;
      if (isInRepair) displayType = equipmentType + ' - In Repair';
      else if (isPendingJob) displayType = equipmentType + ' - Pending Job';
      if (orderNumbersFound.size > 0) {
        orderNumbersFound.forEach(function (normOrder) {
          if (!camerasByOrder[normOrder]) camerasByOrder[normOrder] = [];
          if (!camerasByOrder[normOrder].some(function (c) { return c.barcode === barcode; })) {
            camerasByOrder[normOrder].push({ equipmentType: displayType, barcode: barcode });
          }
        });
      }
    }
    return camerasByOrder;
  } catch (e) {
    Logger.log('processEquipmentSheet: ' + e.message);
    return {};
  }
}

/**
 * Reads Equipment Scheduling Chart for a target date (Camera + Consignor Use Only), LOS ANGELES rows.
 * @param {Date} targetDate
 * @returns {Object} orderNumber (normalized) -> [{ equipmentType, barcode }, ...]
 */
function readEquipmentSchedulingData(targetDate) {
  try {
    const ss = SpreadsheetApp.openById(PREP_BAY_EQUIPMENT_CHART_ID);
    const cameraBodiesLookup = readCameraBodiesOnlyLookup();
    const camerasByOrder = {};
    const cameraSheet = ss.getSheetByName('Camera');
    if (cameraSheet) {
      const cameraData = processEquipmentSheet(cameraSheet, 'Camera', null, targetDate);
      Object.keys(cameraData).forEach(function (orderNumber) {
        if (!camerasByOrder[orderNumber]) camerasByOrder[orderNumber] = [];
        cameraData[orderNumber].forEach(function (cam) {
          if (!camerasByOrder[orderNumber].some(function (c) { return c.barcode === cam.barcode; })) {
            camerasByOrder[orderNumber].push(cam);
          }
        });
      });
    }
    const consignorSheet = ss.getSheetByName('Consignor Use Only');
    if (consignorSheet) {
      const consignorData = processEquipmentSheet(consignorSheet, 'Consignor Use Only', cameraBodiesLookup, targetDate);
      Object.keys(consignorData).forEach(function (orderNumber) {
        if (!camerasByOrder[orderNumber]) camerasByOrder[orderNumber] = [];
        consignorData[orderNumber].forEach(function (cam) {
          if (!camerasByOrder[orderNumber].some(function (c) { return c.barcode === cam.barcode; })) {
            camerasByOrder[orderNumber].push(cam);
          }
        });
      });
    }
    return camerasByOrder;
  } catch (e) {
    Logger.log('readEquipmentSchedulingData: ' + e.message);
    return {};
  }
}

/**
 * Groups Prep Bay rows by order number and builds one job summary per order (prep bays as "1, 2, & 3").
 * @param {Array<Object>} prepBayData - from readPrepBayDataForDate
 * @returns {Array<Object>} [{ jobName, orderNumber, prepBaysDisplay, prepTech, notes, bayNumbers }, ...]
 */
function groupPrepBayByOrder(prepBayData) {
  const byOrder = {};
  prepBayData.forEach(function (a) {
    const norm = String(a.orderNumber || '').replace(/[^0-9]/g, '');
    if (!norm) return;
    if (!byOrder[norm]) {
      byOrder[norm] = { jobName: a.jobName, orderNumber: a.orderNumber, prepTech: a.prepTech, notes: a.notes, bayNumbers: [] };
    }
    byOrder[norm].bayNumbers.push(a.bayNumber);
  });
  const jobs = [];
  Object.keys(byOrder).forEach(function (norm) {
    const j = byOrder[norm];
    j.bayNumbers.sort(function (a, b) { return a - b; });
    let prepBaysDisplay = j.bayNumbers.map(function (b) {
      if (b >= 1 && b <= 19) return String(b);
      if (b === 20) return 'BL 1';
      if (b === 21) return 'BL 2';
      if (b === 22) return 'KTN';
      return String(b);
    }).join(', ');
    if (j.bayNumbers.length >= 2) {
      const last = j.bayNumbers[j.bayNumbers.length - 1];
      const rest = j.bayNumbers.slice(0, -1).join(', ');
      prepBaysDisplay = rest + ' & ' + (last >= 20 ? (last === 20 ? 'BL 1' : last === 21 ? 'BL 2' : 'KTN') : last);
    }
    jobs.push({
      jobName: j.jobName,
      orderNumber: j.orderNumber,
      prepBaysDisplay: prepBaysDisplay,
      marketingAgent: '',
      prepTech: j.prepTech,
      prepNotes: j.notes,
      bayNumbers: j.bayNumbers
    });
  });
  return jobs.sort(function (a, b) {
    const aMin = Math.min.apply(null, a.bayNumbers);
    const bMin = Math.min.apply(null, b.bayNumbers);
    return aMin - bMin;
  });
}

/** Pad a row to column count 9 for consistent setValues */
function padRow(arr) {
  const out = arr.slice();
  while (out.length < 9) out.push('');
  return out.slice(0, 9);
}

/**
 * Builds one job block (v2 layout) as a 2D array for the sheet.
 * Columns: A = label, B = value (or sub-table). Minimal horizontal width.
 *
 * @param {Object} job - { jobName, orderNumber, prepBaysDisplay, marketingAgent, prepTech, prepNotes }
 * @returns {Array<Array>} Rows for this job block (each row length 9)
 */
function buildJobBlockRows(job) {
  const rows = [];

  rows.push(padRow(['Job Name:', job.jobName || '']));
  rows.push(padRow(['Order #:', job.orderNumber || '']));
  rows.push(padRow(['Prep Bay(s):', job.prepBaysDisplay || '']));
  rows.push(padRow(['Marketing Agent:', job.marketingAgent || '']));
  rows.push(padRow(['Prep Tech:', job.prepTech || '']));
  rows.push(padRow(['Prep Notes:', job.prepNotes || '']));
  rows.push(padRow([]));

  // Equipment table header
  rows.push(padRow(['', 'Equipment Name', 'Barcode', 'Pulled?', 'RTR?', 'Serviced for Order?', 'Completion Timestamp']));
  EQUIPMENT_CATEGORIES.forEach(function (cat) {
    rows.push(padRow([cat + ':', '', '', false, false, false, '']));
  });
  rows.push(padRow([]));

  // Sub-rental section
  rows.push(padRow(['', 'Subbed Equipment', 'Quantity', 'Located', 'Locating Agent', 'Quote Rec.', 'Run Sheet Out', 'Packing Slip', 'Notes']));
  rows.push(padRow(['', '', '', false, false, false, '', '', '']));
  rows.push(padRow([]));

  return rows;
}

/**
 * Builds one job block rows with scheduled cameras from Equipment Chart (same layout as buildJobBlockRows).
 * Fills up to 10 equipment rows with camera data; remaining rows keep category labels and empty cells.
 * @param {Object} job - { jobName, orderNumber, prepBaysDisplay, marketingAgent, prepTech, prepNotes }
 * @param {Array<Object>} cameras - [{ equipmentType, barcode }, ...] from readEquipmentSchedulingData
 * @returns {Array<Array>}
 */
function buildJobBlockRowsWithCameras(job, cameras) {
  const rows = [];
  rows.push(padRow(['Job Name:', job.jobName || '']));
  rows.push(padRow(['Order #:', job.orderNumber || '']));
  rows.push(padRow(['Prep Bay(s):', job.prepBaysDisplay || '']));
  rows.push(padRow(['Marketing Agent:', job.marketingAgent || '']));
  rows.push(padRow(['Prep Tech:', job.prepTech || '']));
  rows.push(padRow(['Prep Notes:', job.prepNotes || '']));
  rows.push(padRow([]));
  rows.push(padRow(['', 'Equipment Name', 'Barcode', 'Pulled?', 'RTR?', 'Serviced for Order?', 'Completion Timestamp']));
  const maxEquipmentRows = 10;
  const camList = cameras || [];
  for (let i = 0; i < maxEquipmentRows; i++) {
    const label = EQUIPMENT_CATEGORIES[i] ? EQUIPMENT_CATEGORIES[i] + ':' : '';
    if (i < camList.length) {
      rows.push(padRow([label, camList[i].equipmentType || '', camList[i].barcode || '', false, false, false, '']));
    } else {
      rows.push(padRow([label, '', '', false, false, false, '']));
    }
  }
  rows.push(padRow([]));
  rows.push(padRow(['', 'Subbed Equipment', 'Quantity', 'Located', 'Locating Agent', 'Quote Rec.', 'Run Sheet Out', 'Packing Slip', 'Notes']));
  rows.push(padRow(['', '', '', false, false, false, '', '', '']));
  rows.push(padRow([]));
  return rows;
}

/**
 * Applies formatting to one job block using the given format object (FMT_DEFAULTS).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow - 1-based starting row of the job block
 * @param {Object} fmt - Format options (defaults to FMT_DEFAULTS)
 * @param {string|null} jobHeaderBgOverride - Optional: use this as job header band background (e.g. from Prep Bay column C); if null, use fmt.jobHeaderBg
 */
function applyJobBlockFormatting(sheet, startRow, fmt, jobHeaderBgOverride) {
  if (!fmt) fmt = FMT_DEFAULTS;
  const r = startRow;
  const numCols = 9;
  const jobBg = jobHeaderBgOverride != null && jobHeaderBgOverride !== '' ? jobHeaderBgOverride : fmt.jobHeaderBg;

  // --- Job header (rows 1–6): entire block uses background from Prep Bay column C (order # cell) ---
  sheet.getRange(r, 1, r + 5, numCols).setBackground(jobBg);
  sheet.getRange(r, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor);
  sheet.getRange(r, 2).setFontWeight('bold').setFontSize(fmt.jobNameValueSize).setFontColor(fmt.jobNameValueColor);
  sheet.setRowHeight(r, fmt.rowHeightJobName);

  for (let i = 1; i <= 5; i++) {
    const row = r + i;
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor);
    sheet.getRange(row, 2).setFontSize(fmt.valueSize);
    sheet.setRowHeight(row, fmt.rowHeightLabel);
  }

  sheet.getRange(r + 6, 1, r + 6, numCols).setBorder(null, null, true, null, null, null, fmt.borderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const eqHeaderRow = r + 7;
  sheet.getRange(eqHeaderRow, 1, eqHeaderRow, numCols).setBackground(fmt.tableHeaderBg).setFontColor(fmt.tableHeaderFg).setFontWeight('bold').setFontSize(fmt.tableHeaderSize);
  sheet.setRowHeight(eqHeaderRow, fmt.rowHeightTableHeader);

  for (let i = 0; i < EQUIPMENT_CATEGORIES.length; i++) {
    const row = eqHeaderRow + 1 + i;
    sheet.getRange(row, 1, row, numCols).setBackground(fmt.categoryBg);
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(fmt.valueSize);
    sheet.setRowHeight(row, fmt.rowHeightCategory);
  }

  // --- Subbed Equipment: header same blue + white bold as Equipment Name; data row keeps default ---
  const subHeaderRow = r + 19;
  sheet.getRange(subHeaderRow, 1, subHeaderRow, numCols).setBackground(fmt.tableHeaderBg).setFontColor(fmt.tableHeaderFg).setFontWeight('bold').setFontSize(fmt.tableHeaderSize);
  sheet.setRowHeight(subHeaderRow, fmt.rowHeightTableHeader);
  sheet.setRowHeight(subHeaderRow + 1, fmt.rowHeightCategory);

  // --- Black horizontal line separating this order from the next ---
  const lastRow = r + ROWS_PER_JOB_BLOCK - 1;
  sheet.getRange(lastRow, 1, lastRow, numCols).setBackground('#000000');
}

/**
 * Sets column widths from the format object (FMT_DEFAULTS).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} fmt - Format options (defaults to FMT_DEFAULTS)
 */
function applySchemaColumnWidths(sheet, fmt) {
  if (!fmt) fmt = FMT_DEFAULTS;
  sheet.setColumnWidth(1, fmt.colWidthLabel);
  sheet.setColumnWidth(2, fmt.colWidthValue);
  for (let c = 3; c <= 9; c++) {
    sheet.setColumnWidth(c, 96);
  }
}

/**
 * Writes v2 sample data to "Prep Bay Schema" sheet (two job blocks) with formatting.
 * Run from the Apps Script editor bound to the LA Prep workbook.
 */
function writePrepBaySchemaTest() {
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  let sheet = ss.getSheetByName(PREP_BAY_SCHEMA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PREP_BAY_SCHEMA_SHEET_NAME);
  }

  const job1 = {
    jobName: '"SAMPLE JOB NAME HERE"',
    orderNumber: '123456',
    prepBaysDisplay: '1, 2, & 3',
    marketingAgent: 'Joe (sample agent name)',
    prepTech: 'Bobby (sample prep tech name)',
    prepNotes: 'Preps 3/1 - 3/3  (sample data)'
  };
  const job2 = {
    jobName: '"2nd SAMPLE JOB NAME HERE"',
    orderNumber: '123457',
    prepBaysDisplay: '5',
    marketingAgent: 'Mary (sample agent name)',
    prepTech: 'Pete (sample prep tech name)',
    prepNotes: 'Preps 3/1 - 3/6  (sample data)'
  };

  const allRows = []
    .concat(buildJobBlockRows(job1))
    .concat(buildJobBlockRows(job2));

  const numRows = allRows.length;
  const numCols = 9;
  if (numRows === 0) return;

  sheet.clear();
  sheet.getRange(1, 1, numRows, numCols).setValues(allRows);
  sheet.getRange(1, 1, numRows, numCols).setWrap(true);

  const fmt = FMT_DEFAULTS;
  applySchemaColumnWidths(sheet, fmt);

  const prepBaySheetName = getTodaySheetName(new Date());
  const job1HeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job1.orderNumber);
  const job2HeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job2.orderNumber);

  applyJobBlockFormatting(sheet, 1, fmt, job1HeaderBg);
  applyJobBlockFormatting(sheet, 1 + ROWS_PER_JOB_BLOCK, fmt, job2HeaderBg);

  Logger.log('Prep Bay Schema test written: ' + numRows + ' rows');
}

/**
 * Refreshes all five forecast sheets (Prep Today, Prep Tomorrow, Prep Two Days Out, Prep Three Days Out, Prep Four Days Out)
 * with live data from Prep Bay Assignment and Equipment Scheduling Chart.
 * Jobs are grouped by order number; scheduled cameras for each order are filled from the Equipment Chart.
 */
function refreshPrepForecastSheets() {
  Logger.log('Starting refreshPrepForecastSheets');
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  const fmt = FMT_DEFAULTS;
  const numCols = 9;

  PREP_FORECAST_SHEETS.forEach(function (config) {
    const sheetName = config.name;
    const daysOffset = config.daysOffset;
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + daysOffset);
    const prepBaySheetName = getTodaySheetName(targetDate);

    Logger.log('Processing ' + sheetName + ' (date: ' + prepBaySheetName + ')...');

    const prepBayData = readPrepBayDataForDate(prepBaySheetName);
    const equipmentData = readEquipmentSchedulingData(targetDate);
    const jobs = groupPrepBayByOrder(prepBayData);

    const allRows = [];
    jobs.forEach(function (job) {
      const normOrder = String(job.orderNumber || '').replace(/[^0-9]/g, '');
      const cameras = equipmentData[normOrder] || [];
      allRows.push.apply(allRows, buildJobBlockRowsWithCameras(job, cameras));
    });

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Created sheet: ' + sheetName);
    }

    sheet.clear();
    if (allRows.length > 0) {
      sheet.getRange(1, 1, allRows.length, numCols).setValues(allRows);
      sheet.getRange(1, 1, allRows.length, numCols).setWrap(true);
      applySchemaColumnWidths(sheet, fmt);
      var startRow = 1;
      jobs.forEach(function (job) {
        var jobHeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job.orderNumber);
        applyJobBlockFormatting(sheet, startRow, fmt, jobHeaderBg);
        startRow += ROWS_PER_JOB_BLOCK;
      });
    }

    SpreadsheetApp.flush();
    Logger.log('Done ' + sheetName + ' (' + jobs.length + ' jobs, ' + allRows.length + ' rows).');
  });

  Logger.log('refreshPrepForecastSheets done');
}
