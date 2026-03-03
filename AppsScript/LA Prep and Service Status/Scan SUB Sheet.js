/**
 * Scan SUB Sheet — runs on its own (menu or time-driven trigger every 5–10 min).
 * Scans the SUB workbook for subbed equipment by order and writes the "Sub Equipment Helper"
 * sheet in the LA Prep workbook. Prep refresh (single-day and Initialize) does not run this;
 * they only read from the helper sheet, so refresh stays fast.
 *
 * Trigger: Add a time-driven trigger (Edit > Current project's triggers) for runScanSubSheet
 * every 5–10 minutes if desired.
 *
 * Uses LA_PREP_STATUS_WORKBOOK_ID and SUB_HELPER_SHEET_NAME from Prep Bay Schema Test.js (same project).
 */

const SUB_SHEET_WORKBOOK_ID = '1UUAwABLOAQLt9M4uTa8E6DkdLJSKHIC7s_xZDzLtBQw';
const SUB_SHEET_TEMPLATE_NAME = 'Template - Please Copy to Create Tabs';

// SUB sheet block layout (rows 1–6 are human-only; order number drives the block):
// - Order number: B7, then every 13 rows (B7, B20, B33, …).
// - Header row: one row below order (row 8, 21, 34, …). Data: 10 rows after header (rows 9–18, 22–31, …).
const SUB_BLOCK_FIRST_ROW = 7;   // row containing first order number (B7)
const SUB_BLOCK_ROW_COUNT = 13;  // next order at +13 (e.g. 7 → 20 → 33)
const SUB_BLOCK_DATA_ROWS = 10;  // data rows per block (e.g. 9–18)
const SUB_HEADER_OFFSET = 1;     // header row = order row + 1
const SUB_DATA_OFFSET = 2;      // first data row = order row + 2

// Column layout (header row 8): A=blank, B=Time Needed By, C=QTY, D:E=Requested Equipment, F=Billing Days (unused),
// G=Added By, H=Agent Signoff, I=Called (unused), J=Located, K=Quote Received, L=Run Sheet w/ Shipping, M=Packing Slip, N:P=Notes, Q=Locating Agent (header in Q6; data starts Q9).
const SUB_NUM_COLS = 17;        // A–Q

/**
 * Returns true if sheetName looks like a date (e.g. "Monday 3/2/26") and that date is before today.
 * Used to skip past-date sheets that haven't been hidden yet.
 */
function isPastDateSheetName(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') return false;
  var match = String(sheetName).trim().match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!match) return false;
  var month = parseInt(match[1], 10) - 1;
  var day = parseInt(match[2], 10);
  var year = parseInt(match[3], 10);
  if (year < 100) year += 2000;
  var sheetDate = new Date(year, month, day);
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  sheetDate.setHours(0, 0, 0, 0);
  return sheetDate.getTime() < today.getTime();
}

/**
 * Returns true if the row has strikethrough on Requested Equipment (D or E) in the SUB block (cancelled).
 */
function subBlockRowHasStrikethrough(sheet, dataStartRow, rowOffset) {
  try {
    const row = dataStartRow + rowOffset;
    const cellD = sheet.getRange(row, 4).getTextStyle();
    const cellE = sheet.getRange(row, 5).getTextStyle();
    return !!(cellD && cellD.isStrikethrough() === true) || !!(cellE && cellE.isStrikethrough() === true);
  } catch (err) {
    return false;
  }
}

/**
 * Parses one SUB block data array (10 rows × 17 cols, A–Q) into items.
 * Column mapping: B=Time Needed By, C=QTY, D:E=Requested Equipment, G=Added By, H=Agent Signoff, J=Located, K=Quote Received, L=Run Sheet w/ Shipping, M=Packing Slip, N:P=Notes, Q=Locating Agent.
 */
function parseSubBlockData(data, strikethroughByRow) {
  const items = [];
  const hasCheck = function (v) { return /✓|✔|yes|true|1/i.test(v); };
  const strike = strikethroughByRow || [];
  for (let i = 0; i < (data && data.length ? data.length : 0); i++) {
    const row = data[i] || [];
    const qty = row[2] != null ? String(row[2]).trim() : '';                    // C = QTY
    const reqD = row[3] != null ? String(row[3]).trim() : '';                   // D = Requested Equipment
    const reqE = row[4] != null ? String(row[4]).trim() : '';                   // E = Requested Equipment
    const subbedEquipment = (reqD + ' ' + reqE).trim() || '';
    const addedBy = row[6] != null ? String(row[6]).trim() : '';                 // G = Added By
    const locatingAgent = row[16] != null ? String(row[16]).trim() : '';        // Q = Locating Agent (Q9–Q18)
    const located = row[9] != null ? String(row[9]).trim() : '';                // J = Located
    const quoteReceived = hasCheck(row[10] != null ? String(row[10]).trim() : '');  // K = Quote Received
    const runSheet = hasCheck(row[11] != null ? String(row[11]).trim() : '');      // L = Run Sheet w/ Shipping
    const packingSlip = hasCheck(row[12] != null ? String(row[12]).trim() : '');  // M = Packing Slip
    const notesN = row[13] != null ? String(row[13]).trim() : '';               // N:P = Notes
    const notesO = row[14] != null ? String(row[14]).trim() : '';
    const notesP = row[15] != null ? String(row[15]).trim() : '';
    const notes = [notesN, notesO, notesP].filter(Boolean).join(' ').trim();
    if (!subbedEquipment && !qty && !located && !notes) continue;
    items.push({ addedBy: addedBy, subbedEquipment: subbedEquipment, qty: qty, locatingAgent: locatingAgent, located: located, quoteReceived: quoteReceived, runSheet: runSheet, packingSlip: packingSlip, notes: notes, strikethrough: !!strike[i] });
  }
  return items;
}

/**
 * Scans entire SUB workbook (all non-template, non-hidden sheets, all blocks). Returns map normOrder -> items.
 * When the same quote appears in more than one block, we merge: all items from every matching block are included.
 */
function scanSubWorkbookIntoMap() {
  const map = {};
  let t0 = new Date().getTime();
  try {
    const ss = SpreadsheetApp.openById(SUB_SHEET_WORKBOOK_ID);
    let t1 = new Date().getTime();
    Logger.log('[Scan SUB] open SUB workbook: ' + (t1 - t0) + ' ms');
    const sheets = ss.getSheets();
    Logger.log('[Scan SUB] getSheets() count: ' + sheets.length);
    for (let s = 0; s < sheets.length; s++) {
      const sheetStart = new Date().getTime();
      const sheet = sheets[s];
      const sheetName = sheet.getName();
      if (sheetName === SUB_SHEET_TEMPLATE_NAME) {
        Logger.log('[Scan SUB]   skip sheet (template): ' + sheetName);
        continue;
      }
      if (sheet.isSheetHidden()) {
        Logger.log('[Scan SUB]   skip sheet (hidden): ' + sheetName);
        continue;
      }
      if (isPastDateSheetName(sheetName)) {
        Logger.log('[Scan SUB]   skip sheet (past date): ' + sheetName);
        continue;
      }
      const lastRow = sheet.getLastRow();
      if (lastRow < SUB_BLOCK_FIRST_ROW + SUB_DATA_OFFSET) {
        Logger.log('[Scan SUB]   skip sheet (no data): ' + sheetName);
        continue;
      }
      let blocksWithQuote = 0;
      for (let k = 0; k < 18; k++) {
        const orderRow = SUB_BLOCK_FIRST_ROW + k * SUB_BLOCK_ROW_COUNT;
        if (orderRow + SUB_DATA_OFFSET + SUB_BLOCK_DATA_ROWS - 1 > lastRow) break;
        const quoteCell = sheet.getRange(orderRow, 2).getValue();
        const quoteNorm = String(quoteCell || '').replace(/[^0-9]/g, '');
        if (!quoteNorm) continue;
        blocksWithQuote++;
        const dataStartRow = orderRow + SUB_DATA_OFFSET;
        const data = sheet.getRange(dataStartRow, 1, SUB_BLOCK_DATA_ROWS, SUB_NUM_COLS).getValues();
        const strikethroughByRow = [];
        for (let i = 0; i < data.length; i++) {
          strikethroughByRow.push(subBlockRowHasStrikethrough(sheet, dataStartRow, i));
        }
        const items = parseSubBlockData(data, strikethroughByRow);
        if (items.length > 0) {
          if (!map[quoteNorm]) map[quoteNorm] = [];
          map[quoteNorm].push.apply(map[quoteNorm], items);
        }
      }
      const sheetElapsed = new Date().getTime() - sheetStart;
      Logger.log('[Scan SUB]   sheet "' + sheetName + '": ' + sheetElapsed + ' ms (blocksWithQuote=' + blocksWithQuote + ')');
    }
    const scanTotal = new Date().getTime() - t0;
    Logger.log('[Scan SUB] scanSubWorkbookIntoMap total: ' + scanTotal + ' ms');
  } catch (e) {
    Logger.log('scanSubWorkbookIntoMap: ' + e.message);
  }
  return map;
}

/**
 * Scan SUB Sheet: full scan of SUB workbook, writes "Sub Equipment Helper" sheet in LA Prep workbook.
 * Call from menu (Prep Refresh > Scan SUB Sheet) or from a time-driven trigger (e.g. every 5–10 min).
 * Prep refresh (single-day and Initialize) only reads from the helper sheet and does not run this.
 */
function runScanSubSheet() {
  const runStart = new Date().getTime();
  Logger.log('[Scan SUB] ========== runScanSubSheet started ==========');

  let t0 = new Date().getTime();
  const map = scanSubWorkbookIntoMap();
  const orderCount = Object.keys(map).length;
  Logger.log('[Scan SUB] scan phase total: ' + (new Date().getTime() - t0) + ' ms');

  t0 = new Date().getTime();
  const rows = [['OrderNumber', 'AddedBy', 'SubbedEquipment', 'Qty', 'LocatingAgent', 'Located', 'QuoteReceived', 'RunSheet', 'PackingSlip', 'Notes', 'Strikethrough']];
  Object.keys(map).forEach(function (normOrder) {
    map[normOrder].forEach(function (item) {
      rows.push([normOrder, item.addedBy || '', item.subbedEquipment, item.qty, item.locatingAgent || '', item.located, item.quoteReceived, item.runSheet, item.packingSlip, item.notes, item.strikethrough === true]);
    });
  });
  Logger.log('[Scan SUB] build rows array: ' + (new Date().getTime() - t0) + ' ms (rows=' + rows.length + ')');

  t0 = new Date().getTime();
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  Logger.log('[Scan SUB] open LA Prep workbook: ' + (new Date().getTime() - t0) + ' ms');

  let sheet = ss.getSheetByName(SUB_HELPER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUB_HELPER_SHEET_NAME);
    Logger.log('[Scan SUB] created sheet: ' + SUB_HELPER_SHEET_NAME);
  }

  t0 = new Date().getTime();
  sheet.clear();
  Logger.log('[Scan SUB] sheet.clear(): ' + (new Date().getTime() - t0) + ' ms');

  t0 = new Date().getTime();
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 11).setValues(rows);
  } else {
    sheet.getRange(1, 1, 1, 11).setValues([rows[0]]);
  }
  Logger.log('[Scan SUB] setValues(' + rows.length + 'x11): ' + (new Date().getTime() - t0) + ' ms');

  t0 = new Date().getTime();
  SpreadsheetApp.flush();
  Logger.log('[Scan SUB] flush(): ' + (new Date().getTime() - t0) + ' ms');

  const runTotal = new Date().getTime() - runStart;
  Logger.log('[Scan SUB] ========== done. ' + orderCount + ' orders, ' + (rows.length - 1) + ' items. Total run: ' + runTotal + ' ms ==========');
}
