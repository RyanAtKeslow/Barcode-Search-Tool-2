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

/** Forecast sheet names and workday offset (0 = today, 1 = next workday; weekends and federal holidays skipped) */
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

/** Row count per job block (job header 4 + eq header 1 + categories + sub header 1 + sub row(s) + black 1). Job header: A1:B1, row 2 C2:D2 (Marketing Agent) F2:G2 (DP), row 3 C3:D3 (Prep Tech) F3:G3 (1st AC), row 4 Prep Notes. */
const ROWS_PER_JOB_BLOCK = 18;

/**
 * Job block row names (by column A content; used for clarity and locating rows).
 * Row 1: Job Name (A1:B1). Row 2: Order # (A2:B2), Marketing Agent (C2:D2), DP (F2:G2). Row 3: Prep Bay(s) (A3:B3), Prep Tech (C3:D3), 1st AC (F3:G3). Row 4: Prep Notes (A4:B4).
 * Equipment Header (A empty) | equipment rows | Locating Agent | sub data row(s) | black bar.
 * Checkboxes: Equipment = D,E,F. Sub = D,E,F,G.
 */

/** Default formatting (used when Settings sheet is missing or a value is blank) */
const FMT_DEFAULTS = {
  jobHeaderBg: '#e8f0fe',
  jobNameValueSize: 28,
  jobNameValueColor: '#1a73e8',
  labelColor: '#000000',
  labelSize: 11,
  valueSize: 11,
  valueColor: '#000000',
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

/**
 * Adds the "Prep Refresh" menu when the workbook is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Prep Refresh')
    .addItem('Initialize Prep Refresh', 'refreshPrepForecastSheets')
    .addSeparator()
    .addItem('Refresh Today', 'refreshPrepToday')
    .addItem('Refresh Tomorrow', 'refreshPrepTomorrow')
    .addItem('Refresh Two Days Out', 'refreshPrepTwoDaysOut')
    .addItem('Refresh Three Days Out', 'refreshPrepThreeDaysOut')
    .addItem('Refresh Four Days Out', 'refreshPrepFourDaysOut')
    .addSeparator()
    .addItem('Job Block Test', 'runJobBlockTest')
    .addToUi();
}

/**
 * Normalizes equipment items for the dynamic block builder.
 * Accepts either legacy cameras array [{ equipmentType, barcode }] (treated as Cameras)
 * or full format [{ category, equipmentType, barcode }]. Returns array of { category, equipmentType, barcode }.
 * @param {Array<Object>} equipmentList - Cameras or full equipment list
 * @param {string} defaultCategory - Category when item has no category (e.g. 'Cameras')
 * @returns {Array<{category: string, equipmentType: string, barcode: string}>}
 */
function normalizeEquipmentByCategory(equipmentList, defaultCategory) {
  const cat = defaultCategory || 'Cameras';
  return (equipmentList || []).map(function (item) {
    return {
      category: item.category && EQUIPMENT_CATEGORIES.indexOf(item.category) >= 0 ? item.category : cat,
      equipmentType: item.equipmentType || item.name || '',
      barcode: item.barcode || ''
    };
  });
}

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

/**
 * Whether the date is a US federal / company-observed holiday (Prep Bay closed).
 * Same set as Tomorrow Double Check / Camera Forecast: New Year's, Memorial Day,
 * Independence Day, Labor Day, Thanksgiving, Day after Thanksgiving, Christmas,
 * Christmas week closure (Dec 26–31).
 * @param {Date} date
 * @returns {boolean}
 */
function isUSFederalHoliday(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const dayOfWeek = date.getDay();

  function getNthWeekday(y, m, weekday, n) {
    const first = new Date(y, m, 1);
    const firstWd = first.getDay();
    let d = 1 + (weekday - firstWd + 7) % 7;
    d += (n - 1) * 7;
    return new Date(y, m, d);
  }
  function getLastWeekday(y, m, weekday) {
    const last = new Date(y, m + 1, 0);
    const lastWd = last.getDay();
    const d = last.getDate() - ((lastWd - weekday + 7) % 7);
    return new Date(y, m, d);
  }

  if (month === 0 && day === 1) return true;
  if (month === 0 && day === 2 && dayOfWeek === 1) return true;
  if (month === 11 && day === 31 && dayOfWeek === 5) return true;

  const memorialDay = getLastWeekday(year, 4, 1);
  if (month === memorialDay.getMonth() && day === memorialDay.getDate()) return true;

  if (month === 6 && day === 4) return true;
  if (month === 6 && day === 3 && dayOfWeek === 5) return true;
  if (month === 6 && day === 5 && dayOfWeek === 1) return true;

  const laborDay = getNthWeekday(year, 8, 1, 1);
  if (month === laborDay.getMonth() && day === laborDay.getDate()) return true;

  const thanksgiving = getNthWeekday(year, 10, 4, 4);
  if (month === thanksgiving.getMonth() && day === thanksgiving.getDate()) return true;
  const dayAfterThanksgiving = new Date(thanksgiving);
  dayAfterThanksgiving.setDate(dayAfterThanksgiving.getDate() + 1);
  if (month === dayAfterThanksgiving.getMonth() && day === dayAfterThanksgiving.getDate()) return true;

  if (month === 11 && day === 25) return true;
  if (month === 11 && day === 24 && dayOfWeek === 5) return true;
  if (month === 11 && day === 26 && dayOfWeek === 1) return true;
  if (month === 11 && day >= 26 && day <= 31) return true;

  return false;
}

/**
 * Next business day (Monday–Friday, skip weekends and federal holidays).
 * @param {Date} date
 * @returns {Date}
 */
function getNextBusinessDay(date) {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const dow = date.getDay();
  if (dow === 5) next.setDate(date.getDate() + 3);
  else if (dow === 6) next.setDate(date.getDate() + 2);
  else if (dow === 0) next.setDate(date.getDate() + 1);
  else next.setDate(date.getDate() + 1);
  while (isUSFederalHoliday(next) || next.getDay() === 0 || next.getDay() === 6) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

/**
 * Date for a forecast sheet by workday offset (skip weekends and holidays).
 * Offset 0 = today; 1 = next workday; 2 = 2nd workday out; etc.
 * Prep Bay Assignment only has weekdays, so "Prep Four Days Out" on Tuesday = Monday (skip Sat/Sun).
 * @param {Date} startDate - usually today
 * @param {number} workdayOffset - 0, 1, 2, 3, or 4
 * @returns {Date}
 */
function getDateForForecastOffset(startDate, workdayOffset) {
  const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  if (workdayOffset === 0) return d;
  for (let i = 0; i < workdayOffset; i++) {
    const next = getNextBusinessDay(d);
    d.setTime(next.getTime());
  }
  return d;
}

/** Returns the day-of-week name in all caps for a date (e.g. "WEDNESDAY"). */
function getDayNameAllCaps(date) {
  const names = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
  return names[date.getDay()] || '';
}

/** Returns the row 1 title for prep sheets: day name + date (e.g. "WEDNESDAY, 2/25/2026"). */
function getPrepSheetTitle(date) {
  const d = date || new Date();
  return getDayNameAllCaps(d) + ', ' + (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
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
 * @returns {Array<Object>} [{ bayNumber, bayName, jobName, orderNumber, marketingAgent, firstAC, cinematographer, prepTech, notes }, ...]
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
      const marketingAgent = row[3] ? String(row[3]).trim() : '';
      const firstAC = row[5] ? String(row[5]).trim() : '';
      const cinematographer = row[6] ? String(row[6]).trim() : '';
      const prepTech = row[7] ? String(row[7]).trim() : '';
      const notes = row[8] ? String(row[8]).trim() : '';
      if (!bay || bay.toUpperCase() === 'BAY' || !jobName) continue;
      const bayNumber = normalizeBayNumber(bay);
      if (bayNumber == null) continue;
      out.push({ bayNumber: bayNumber, bayName: bay, jobName: jobName, orderNumber: orderNumber, marketingAgent: marketingAgent, firstAC: firstAC, cinematographer: cinematographer, prepTech: prepTech, notes: notes });
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
      byOrder[norm] = { jobName: a.jobName, orderNumber: a.orderNumber, marketingAgent: a.marketingAgent, firstAC: a.firstAC, cinematographer: a.cinematographer, prepTech: a.prepTech, notes: a.notes, bayNumbers: [] };
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
      marketingAgent: j.marketingAgent || '',
      firstAC: j.firstAC || '',
      cinematographer: j.cinematographer || '',
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

/** Pad a row to column count 10 (A–J) for consistent setValues */
function padRow(arr) {
  const out = arr.slice();
  while (out.length < 10) out.push('');
  return out.slice(0, 10);
}

/**
 * Builds one job block (v2 layout) as a 2D array for the sheet.
 * Columns: A = label, B = value (or sub-table). Minimal horizontal width.
 *
 * @param {Object} job - { jobName, orderNumber, prepBaysDisplay, marketingAgent, firstAC, cinematographer, prepTech, prepNotes }
 * @returns {Array<Array>} Rows for this job block (each row length 10)
 */
function buildJobBlockRows(job) {
  const rows = [];

  rows.push(padRow(['Job Name:', job.jobName || '']));
  rows.push(padRow(['Order #:', job.orderNumber || '', 'Marketing Agent:', job.marketingAgent || '', '', 'DP:', job.cinematographer || '']));
  rows.push(padRow(['Prep Bay(s):', job.prepBaysDisplay || '', 'Prep Tech:', job.prepTech || '', '', '1st AC:', job.firstAC || '']));
  rows.push(padRow(['Prep Notes:', job.prepNotes || '']));

  // Equipment table header (no blank row before)
  rows.push(padRow(['', 'Equipment Name', 'Barcode', 'Pulled?', 'RTR?', 'Serviced for Order?', 'Completion Timestamp']));
  EQUIPMENT_CATEGORIES.forEach(function (cat) {
    rows.push(padRow([cat + ':', '', '', false, false, false, '']));
  });

  // Sub-rental section (no blank before or after). Sub data row: checkboxes D,E,F,G (Located, Quote Received, Run Sheet Out, Packing Slip).
  rows.push(padRow(['Locating Agent', 'Subbed Equipment', 'Quantity', 'Located', 'Quote Received', 'Run Sheet Out', 'Packing Slip', 'Notes', '']));
  rows.push(padRow(['', '', '', false, false, false, false, '', ''])); // one sub data row for now; will grow when Sub Sheet is wired
  rows.push(padRow([])); // black separator row (formatted in applyJobBlockFormatting)

  return rows;
}

/**
 * Builds equipment block rows (between Equipment Name header and Locating Agent header).
 * One row per equipment item; category label in column A only on the first row of each category.
 * Categories with no items get one row with label and empty B/C. Order follows EQUIPMENT_CATEGORIES.
 * @param {Array<{category: string, equipmentType: string, barcode: string}>} equipmentByCategory - from normalizeEquipmentByCategory
 * @returns {Array<Array>} Rows (each length 10 via padRow)
 */
function buildEquipmentBlockRows(equipmentByCategory) {
  const rows = [];
  const byCat = {};
  (equipmentByCategory || []).forEach(function (item) {
    const c = item.category;
    if (!byCat[c]) byCat[c] = [];
    byCat[c].push(item);
  });
  EQUIPMENT_CATEGORIES.forEach(function (cat) {
    const items = byCat[cat] || [];
    const n = Math.max(1, items.length);
    for (let i = 0; i < n; i++) {
      const label = i === 0 ? cat + ':' : '';
      const name = i < items.length ? (items[i].equipmentType || '') : '';
      const barcode = i < items.length ? (items[i].barcode || '') : '';
      rows.push(padRow([label, name, barcode, false, false, false, '']));
    }
  });
  return rows;
}

/**
 * Builds one job block with dynamic equipment rows. Equipment list can be legacy cameras
 * [{ equipmentType, barcode }] (all treated as Cameras) or [{ category, equipmentType, barcode }].
 * Column A category labels (Cameras:, Lenses:, ...) appear only on the first row of each category;
 * extra equipment in a category pushes the next category header down.
 * @param {Object} job - { jobName, orderNumber, prepBaysDisplay, marketingAgent, firstAC, cinematographer, prepTech, prepNotes }
 * @param {Array<Object>} equipmentList - Cameras or full equipment; normalized via normalizeEquipmentByCategory
 * @returns {Array<Array>}
 */
function buildJobBlockRowsWithCameras(job, equipmentList) {
  const rows = [];
  rows.push(padRow(['Job Name:', job.jobName || '']));
  rows.push(padRow(['Order #:', job.orderNumber || '', 'Marketing Agent:', job.marketingAgent || '', '', 'DP:', job.cinematographer || '']));
  rows.push(padRow(['Prep Bay(s):', job.prepBaysDisplay || '', 'Prep Tech:', job.prepTech || '', '', '1st AC:', job.firstAC || '']));
  rows.push(padRow(['Prep Notes:', job.prepNotes || '']));

  rows.push(padRow(['', 'Equipment Name', 'Barcode', 'Pulled?', 'RTR?', 'Serviced for Order?', 'Completion Timestamp']));
  const normalized = normalizeEquipmentByCategory(equipmentList, 'Cameras');
  rows.push.apply(rows, buildEquipmentBlockRows(normalized));

  rows.push(padRow(['Locating Agent', 'Subbed Equipment', 'Quantity', 'Located', 'Quote Received', 'Run Sheet Out', 'Packing Slip', 'Notes', '']));
  rows.push(padRow(['', '', '', false, false, false, false, '', ''])); // sub data row(s); D,E,F,G = checkboxes; will grow when Sub Sheet is wired
  rows.push(padRow([])); // black separator row (formatted in applyJobBlockFormatting)
  return rows;
}

/**
 * Applies formatting to one job block using the given format object (FMT_DEFAULTS).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow - 1-based starting row of the job block
 * @param {Object} fmt - Format options (defaults to FMT_DEFAULTS)
 * @param {string|null} jobHeaderBgOverride - Optional: use this as job header band background (e.g. from Prep Bay column C); if null, use fmt.jobHeaderBg
 * @param {number} [blockRowCount] - Optional: number of rows in this block (for variable-height blocks); if omitted, uses ROWS_PER_JOB_BLOCK
 */
function applyJobBlockFormatting(sheet, startRow, fmt, jobHeaderBgOverride, blockRowCount) {
  if (!fmt) fmt = FMT_DEFAULTS;
  const r = startRow;
  const numCols = 10;
  const jobBg = jobHeaderBgOverride != null && jobHeaderBgOverride !== '' ? jobHeaderBgOverride : fmt.jobHeaderBg;
  const blockEndRow = blockRowCount ? r + blockRowCount - 1 : r + ROWS_PER_JOB_BLOCK - 1;

  // Clear all borders in this block first. getRange 4-arg = (row, column, numRows, numColumns).
  const blockNumRows = blockEndRow - r + 1;
  sheet.getRange(r, 1, blockNumRows, numCols).setBorder(false, false, false, false, false, false, null, null);

  // --- Job header (rows 1–4): A1:B3 left as-is; Marketing Agent / Prep Tech in C2:D3; DP / 1st AC in F2:G3; row 4 Prep Notes ---
  sheet.getRange(r, 1, 4, numCols).setBackground(jobBg);
  sheet.getRange(r, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor);
  sheet.getRange(r, 2).setFontWeight('bold').setFontSize(fmt.jobNameValueSize).setFontColor(fmt.jobNameValueColor).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.setRowHeight(r, fmt.rowHeightJobName);

  // Row 2: Order # (A2:B2), Marketing Agent (C2:D2, D overflow), DP (F2:G2, G overflow); B2 center justify
  sheet.getRange(r + 1, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000');
  sheet.getRange(r + 1, 2).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setHorizontalAlignment('center');
  sheet.getRange(r + 1, 3).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000');
  sheet.getRange(r + 1, 4).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.getRange(r + 1, 6).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000').setHorizontalAlignment('center');
  sheet.getRange(r + 1, 7).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.setRowHeight(r + 1, fmt.rowHeightLabel);

  // Row 3: Prep Bay(s) (A3:B3), Prep Tech (C3:D3, D overflow), 1st AC (F3:G3, G overflow); B3 center justify
  sheet.getRange(r + 2, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000');
  sheet.getRange(r + 2, 2).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setHorizontalAlignment('center');
  sheet.getRange(r + 2, 3).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000');
  sheet.getRange(r + 2, 4).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.getRange(r + 2, 6).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000').setHorizontalAlignment('center');
  sheet.getRange(r + 2, 7).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.setRowHeight(r + 2, fmt.rowHeightLabel);

  // Row 4: Prep Notes (A4:B4, B overflow)
  sheet.getRange(r + 3, 1).setFontWeight('bold').setFontSize(fmt.labelSize).setFontColor(fmt.labelColor || '#000000');
  sheet.getRange(r + 3, 2).setFontWeight('bold').setFontSize(18).setFontColor(fmt.valueColor || '#000000').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.setRowHeight(r + 3, fmt.rowHeightLabel);

  // Only border in the block: divider under Prep Notes (row r+3).
  sheet.getRange(r + 3, 1, 1, numCols).setBorder(null, null, true, null, null, null, fmt.borderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // --- Equipment header (no blank row before); column A empty; white bold text ---
  const eqHeaderRow = r + 4;
  sheet.getRange(eqHeaderRow, 1, 1, numCols).setBackground(fmt.tableHeaderBg).setFontColor(fmt.tableHeaderFg).setFontWeight('bold').setFontSize(fmt.tableHeaderSize);
  sheet.setRowHeight(eqHeaderRow, fmt.rowHeightTableHeader);

  // Find Locating Agent / Subbed Equipment header row (column A = "Locating Agent") so equipment block length is dynamic
  const searchEndRow = Math.min(blockEndRow, sheet.getLastRow());
  const numRowsToSearch = searchEndRow - eqHeaderRow;
  const colAVals = numRowsToSearch > 0 ? sheet.getRange(eqHeaderRow + 1, 1, numRowsToSearch, 1).getValues() : [];
  // Fallback: Locating Agent is 2 rows above black bar (sub data, black)
  let subHeaderRow = blockEndRow - 2;
  for (let i = 0; i < colAVals.length; i++) {
    if (String(colAVals[i][0]).trim() === 'Locating Agent') {
      subHeaderRow = eqHeaderRow + 1 + i;
      break;
    }
  }
  const equipmentBlockEndRow = subHeaderRow - 1;
  const numEquipmentBlockRows = equipmentBlockEndRow - eqHeaderRow;

  // All row indices below are block-relative (derived from startRow r). Changing job block start row will move checkboxes with the block.
  const equipmentDataFirstRow = eqHeaderRow + 1;
  const equipmentDataLastRow = equipmentBlockEndRow;

  // Format all equipment data rows (dynamic count): background, black font; bold column A only for category label rows (e.g. "Cameras:")
  for (let row = equipmentDataFirstRow; row <= equipmentDataLastRow; row++) {
    sheet.getRange(row, 1, 1, numCols).setBackground(jobBg).setFontColor('#000000');
    sheet.setRowHeight(row, fmt.rowHeightCategory);
  }
  const aVals = numEquipmentBlockRows > 0 ? sheet.getRange(equipmentDataFirstRow, 1, numEquipmentBlockRows, 1).getValues() : [];
  for (let i = 0; i < aVals.length; i++) {
    const val = String(aVals[i][0]).trim();
    if (val && val.slice(-1) === ':') {
      sheet.getRange(equipmentDataFirstRow + i, 1).setFontWeight('bold').setFontSize(fmt.valueSize);
    } else {
      sheet.getRange(equipmentDataFirstRow + i, 1).setFontWeight('normal').setFontSize(fmt.valueSize);
    }
  }
  // Checkboxes: equipment cols D,E,F — range tied to this block (equipmentDataFirstRow..equipmentDataLastRow). getRange(row, col, numRows, numCols).
  if (numEquipmentBlockRows > 0) {
    sheet.getRange(equipmentDataFirstRow, 4, numEquipmentBlockRows, 3).insertCheckboxes();
    sheet.getRange(equipmentDataFirstRow, 1, numEquipmentBlockRows, numCols).setFontColor('#000000');
  }

  // --- Subbed Equipment: header row (Locating Agent); then sub data row(s) with checkboxes D,E,F,G ---
  sheet.getRange(subHeaderRow, 1, 1, numCols).setBackground(fmt.tableHeaderBg).setFontColor(fmt.tableHeaderFg).setFontWeight('bold').setFontSize(fmt.tableHeaderSize);
  sheet.setRowHeight(subHeaderRow, fmt.rowHeightTableHeader);
  const numSubRows = blockEndRow - subHeaderRow - 1; // between Locating Agent and black bar; will grow when Sub Sheet is wired
  const subDataFirstRow = subHeaderRow + 1;
  const subDataLastRow = blockEndRow - 1;
  for (let row = subDataFirstRow; row <= subDataLastRow; row++) {
    sheet.setRowHeight(row, fmt.rowHeightCategory);
  }
  // Checkboxes: sub cols D,E,F,G — range tied to this block (subDataFirstRow..subDataLastRow). getRange(row, col, numRows, numCols).
  if (numSubRows > 0) {
    sheet.getRange(subDataFirstRow, 4, numSubRows, 4).insertCheckboxes();
  }

  // --- Black horizontal separator at end of each job block (one row only; getRange 4-arg = row, column, numRows, numColumns) ---
  sheet.getRange(blockEndRow, 1, 1, numCols).setBackground('#000000');
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
  for (let c = 3; c <= 10; c++) {
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
    firstAC: 'Sample 1st AC',
    cinematographer: 'Sample DP',
    prepTech: 'Bobby (sample prep tech name)',
    prepNotes: 'Preps 3/1 - 3/3  (sample data)'
  };
  const job2 = {
    jobName: '"2nd SAMPLE JOB NAME HERE"',
    orderNumber: '123457',
    prepBaysDisplay: '5',
    marketingAgent: 'Mary (sample agent name)',
    firstAC: 'Sample 1st AC',
    cinematographer: 'Sample DP',
    prepTech: 'Pete (sample prep tech name)',
    prepNotes: 'Preps 3/1 - 3/6  (sample data)'
  };

  const titleRow = padRow([getPrepSheetTitle(new Date())]);
  const allRows = [titleRow]
    .concat(buildJobBlockRows(job1))
    .concat(buildJobBlockRows(job2));

  const numRows = allRows.length;
  const numCols = 10;
  if (numRows === 0) return;

  sheet.clear();
  sheet.getRange(1, 1, numRows, numCols).setValues(allRows);
  if (numRows > 1) sheet.getRange(2, 1, numRows, numCols).setWrap(true);

  const fmt = FMT_DEFAULTS;
  applySchemaColumnWidths(sheet, fmt);

  // Row 1: A1 = day name (bold, 32, black, overflow, no background). Set overflow after bulk setWrap so it sticks.
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(32).setFontColor('#000000').setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  const prepBaySheetName = getTodaySheetName(new Date());
  const job1HeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job1.orderNumber);
  const job2HeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job2.orderNumber);

  applyJobBlockFormatting(sheet, 2, fmt, job1HeaderBg);
  applyJobBlockFormatting(sheet, 2 + ROWS_PER_JOB_BLOCK, fmt, job2HeaderBg);

  Logger.log('Prep Bay Schema test written: ' + numRows + ' rows');
}

/**
 * Refreshes a single forecast sheet by workday offset (0 = today, 1 = next workday, etc.).
 * Used by refreshPrepForecastSheets and by the per-sheet menu items (Refresh Today, Refresh Tomorrow, ...).
 * @param {string} sheetName - e.g. "Prep Today", "Prep Tomorrow"
 * @param {number} workdayOffset - 0, 1, 2, 3, or 4 (skip weekends and holidays)
 */
function refreshSinglePrepForecastSheet(sheetName, workdayOffset) {
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  const fmt = FMT_DEFAULTS;
  const numCols = 10;
  const today = new Date();
  const targetDate = getDateForForecastOffset(today, workdayOffset);
  const prepBaySheetName = getTodaySheetName(targetDate);

  Logger.log('Processing ' + sheetName + ' (date: ' + prepBaySheetName + ')...');

  const prepBayData = readPrepBayDataForDate(prepBaySheetName);
  Logger.log('  Prep Bay "' + prepBaySheetName + '": ' + (prepBayData ? prepBayData.length : 0) + ' rows read');

  const equipmentData = readEquipmentSchedulingData(targetDate);
  const equipmentOrderCount = equipmentData ? Object.keys(equipmentData).length : 0;
  Logger.log('  Equipment data: ' + equipmentOrderCount + ' orders with scheduled equipment');

  const jobs = groupPrepBayByOrder(prepBayData);
  Logger.log('  Grouped into ' + jobs.length + ' jobs');

  const allRows = [padRow([getPrepSheetTitle(targetDate)])];
  const blockRowCounts = [];
  jobs.forEach(function (job) {
    const normOrder = String(job.orderNumber || '').replace(/[^0-9]/g, '');
    const cameras = equipmentData[normOrder] || [];
    const blockRows = buildJobBlockRowsWithCameras(job, cameras);
    blockRowCounts.push(blockRows.length);
    allRows.push.apply(allRows, blockRows);
  });

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('  Created sheet: ' + sheetName);
  }

  sheet.clear();
  if (allRows.length > 0) {
    Logger.log('  Writing ' + allRows.length + ' rows, formatting ' + jobs.length + ' job blocks');
    sheet.getRange(1, 1, allRows.length, numCols).setValues(allRows);
    if (allRows.length > 1) sheet.getRange(2, 1, allRows.length, numCols).setWrap(true);
    applySchemaColumnWidths(sheet, fmt);
    // Row 1: A1 = day name (bold, 32, black, overflow, no background). Set overflow after bulk setWrap so it sticks.
    sheet.getRange(1, 1).setFontWeight('bold').setFontSize(32).setFontColor('#000000').setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    var startRow = 2;
    for (var j = 0; j < jobs.length; j++) {
      var jobHeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, jobs[j].orderNumber);
      applyJobBlockFormatting(sheet, startRow, fmt, jobHeaderBg, blockRowCounts[j]);
      Logger.log('    Block ' + (j + 1) + ': order ' + (jobs[j].orderNumber || '') + ', rows ' + blockRowCounts[j] + ', startRow ' + startRow);
      startRow += blockRowCounts[j];
    }
  } else {
    Logger.log('  No jobs for this date; sheet cleared');
  }

  var maxRow = sheet.getMaxRows();
  var lastDataRow = allRows.length || 0;
  if (lastDataRow < maxRow) {
    var trailingRows = maxRow - lastDataRow;
    var trailing = sheet.getRange(lastDataRow + 1, 1, trailingRows, numCols);
    trailing.clearContent().clearFormat().clearDataValidations();
  }

  SpreadsheetApp.flush();
  Logger.log('Done ' + sheetName + ' (' + jobs.length + ' jobs, ' + allRows.length + ' rows).');
}

/** Refreshes only the Prep Today sheet. */
function refreshPrepToday() {
  refreshSinglePrepForecastSheet('Prep Today', 0);
}

/** Refreshes only the Prep Tomorrow sheet. */
function refreshPrepTomorrow() {
  refreshSinglePrepForecastSheet('Prep Tomorrow', 1);
}

/** Refreshes only the Prep Two Days Out sheet. */
function refreshPrepTwoDaysOut() {
  refreshSinglePrepForecastSheet('Prep Two Days Out', 2);
}

/** Refreshes only the Prep Three Days Out sheet. */
function refreshPrepThreeDaysOut() {
  refreshSinglePrepForecastSheet('Prep Three Days Out', 3);
}

/** Refreshes only the Prep Four Days Out sheet. */
function refreshPrepFourDaysOut() {
  refreshSinglePrepForecastSheet('Prep Four Days Out', 4);
}

/**
 * Job Block Test: writes a single job block to the "Job Block Test" sheet starting at A1,
 * using real data for today and the first job (first prep bay) from Prep Bay Assignment.
 * Creates the sheet if it doesn't exist.
 */
function runJobBlockTest() {
  const JOB_BLOCK_TEST_SHEET_NAME = 'Job Block Test';
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  let sheet = ss.getSheetByName(JOB_BLOCK_TEST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(JOB_BLOCK_TEST_SHEET_NAME);
    Logger.log('Created sheet: ' + JOB_BLOCK_TEST_SHEET_NAME);
  }

  const fmt = FMT_DEFAULTS;
  const numCols = 10;
  const today = new Date();
  const prepBaySheetName = getTodaySheetName(today);

  const prepBayData = readPrepBayDataForDate(prepBaySheetName);
  const jobs = groupPrepBayByOrder(prepBayData);

  if (!jobs || jobs.length === 0) {
    sheet.clear();
    Logger.log('Job Block Test: No jobs for today (' + prepBaySheetName + '); sheet cleared.');
    SpreadsheetApp.flush();
    return;
  }

  const job = jobs[0];
  const normOrder = String(job.orderNumber || '').replace(/[^0-9]/g, '');
  const equipmentData = readEquipmentSchedulingData(today);
  const cameras = equipmentData[normOrder] || [];

  const blockRows = buildJobBlockRowsWithCameras(job, cameras);
  const numRows = blockRows.length;

  sheet.clear();
  sheet.getRange(1, 1, numRows, numCols).setValues(blockRows);
  if (numRows > 1) sheet.getRange(2, 1, numRows, numCols).setWrap(true);

  applySchemaColumnWidths(sheet, fmt);

  const jobHeaderBg = getOrderNumberBackgroundFromPrepBay(prepBaySheetName, job.orderNumber);
  applyJobBlockFormatting(sheet, 1, fmt, jobHeaderBg, numRows);

  SpreadsheetApp.flush();
  Logger.log('Job Block Test: wrote 1 job block (' + numRows + ' rows) for order ' + (job.orderNumber || '') + ' from ' + prepBaySheetName);
}

/**
 * Refreshes all five forecast sheets (Prep Today, Prep Tomorrow, Prep Two Days Out, Prep Three Days Out, Prep Four Days Out)
 * with live data from Prep Bay Assignment and Equipment Scheduling Chart.
 */
function refreshPrepForecastSheets() {
  Logger.log('Starting refreshPrepForecastSheets');
  PREP_FORECAST_SHEETS.forEach(function (config) {
    refreshSinglePrepForecastSheet(config.name, config.daysOffset);
  });
  Logger.log('refreshPrepForecastSheets done');
}
