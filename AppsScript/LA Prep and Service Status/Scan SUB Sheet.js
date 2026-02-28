/**
 * Scan SUB Sheet — runs on its own (menu or time-driven trigger every 5–10 min).
 * Scans the SUB workbook for subbed equipment by order and writes the "Sub Equipment Helper"
 * sheet in the LA Prep workbook. Prep refresh (single-day and Initialize) does not run this;
 * they only read from the helper sheet, so refresh stays fast.
 *
 * Trigger: Add a time-driven trigger (Edit > Current project's triggers) for runScanSubSheet
 * every 5–10 minutes if desired.
 */

const LA_PREP_STATUS_WORKBOOK_ID = '1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0';
const SUB_SHEET_WORKBOOK_ID = '1UUAwABLOAQLt9M4uTa8E6DkdLJSKHIC7s_xZDzLtBQw';
const SUB_SHEET_TEMPLATE_NAME = 'Template - Please Copy to Create Tabs';
const SUB_BLOCK_FIRST_ROW = 6;
const SUB_BLOCK_ROW_COUNT = 13;
const SUB_BLOCK_DATA_ROWS = 10;
const SUB_HELPER_SHEET_NAME = 'Sub Equipment Helper';

/**
 * Parses one SUB block data array (10 rows × 16 cols) into items. Same schema as SUB block rows 4–13.
 */
function parseSubBlockData(data) {
  const items = [];
  const hasCheck = function (v) { return /✓|✔|yes|true|1/i.test(v); };
  for (let i = 0; i < (data && data.length ? data.length : 0); i++) {
    const row = data[i] || [];
    const qty = row[2] != null ? String(row[2]).trim() : '';
    const equipD = row[3] != null ? String(row[3]).trim() : '';
    const equipE = row[4] != null ? String(row[4]).trim() : '';
    const subbedEquipment = (equipD + ' ' + equipE).trim() || '';
    const located = row[9] != null ? String(row[9]).trim() : '';
    const kVal = row[10] != null ? String(row[10]).trim() : '';
    const lVal = row[11] != null ? String(row[11]).trim() : '';
    const mVal = row[12] != null ? String(row[12]).trim() : '';
    const quoteReceived = hasCheck(kVal);
    const runSheet = hasCheck(lVal);
    const packingSlip = hasCheck(mVal);
    const notesD = row[13] != null ? String(row[13]).trim() : '';
    const notesE = row[14] != null ? String(row[14]).trim() : '';
    const notesF = row[15] != null ? String(row[15]).trim() : '';
    const notes = [notesD, notesE, notesF].filter(Boolean).join(' ').trim();
    if (!subbedEquipment && !qty && !located && !notes) continue;
    items.push({ subbedEquipment: subbedEquipment, qty: qty, located: located, quoteReceived: quoteReceived, runSheet: runSheet, packingSlip: packingSlip, notes: notes });
  }
  return items;
}

/**
 * Returns true if the row should be skipped: Requested Equipment (D or E) has strikethrough (cancelled).
 */
function subBlockRowHasStrikethrough(sheet, dataStartRow, rowOffset) {
  try {
    const row = dataStartRow + rowOffset;
    const cellD = sheet.getRange(row, 4).getTextStyle();
    const cellE = sheet.getRange(row, 5).getTextStyle();
    const d = cellD && cellD.isStrikethrough() === true;
    const e = cellE && cellE.isStrikethrough() === true;
    return !!(d || e);
  } catch (err) {
    return false;
  }
}

/**
 * Scans entire SUB workbook (all non-template, non-hidden sheets, all blocks). Returns map normOrder -> items.
 * When the same quote appears in more than one block, we merge: all items from every matching block are included.
 */
function scanSubWorkbookIntoMap() {
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(SUB_SHEET_WORKBOOK_ID);
    const sheets = ss.getSheets();
    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      if (sheet.getName() === SUB_SHEET_TEMPLATE_NAME) continue;
      if (sheet.isSheetHidden()) continue;
      const lastRow = sheet.getLastRow();
      if (lastRow < SUB_BLOCK_FIRST_ROW + 1) continue;
      for (let k = 0; k < 18; k++) {
        const startRow = SUB_BLOCK_FIRST_ROW + k * SUB_BLOCK_ROW_COUNT;
        if (startRow + SUB_BLOCK_ROW_COUNT - 1 > lastRow) break;
        const quoteCell = sheet.getRange(startRow + 1, 2).getValue();
        const quoteNorm = String(quoteCell || '').replace(/[^0-9]/g, '');
        if (!quoteNorm) continue;
        const dataStartRow = startRow + 3;
        const numDataRows = SUB_BLOCK_DATA_ROWS;
        const numCols = 16;
        const data = sheet.getRange(dataStartRow, 1, numDataRows, numCols).getValues();
        const filtered = [];
        for (let i = 0; i < data.length; i++) {
          if (!subBlockRowHasStrikethrough(sheet, dataStartRow, i)) filtered.push(data[i]);
        }
        const items = parseSubBlockData(filtered);
        if (items.length > 0) {
          if (!map[quoteNorm]) map[quoteNorm] = [];
          map[quoteNorm].push.apply(map[quoteNorm], items);
        }
      }
    }
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
  Logger.log('Scan SUB Sheet: starting full scan...');
  const map = scanSubWorkbookIntoMap();
  const orderCount = Object.keys(map).length;
  const rows = [['OrderNumber', 'SubbedEquipment', 'Qty', 'Located', 'QuoteReceived', 'RunSheet', 'PackingSlip', 'Notes']];
  Object.keys(map).forEach(function (normOrder) {
    map[normOrder].forEach(function (item) {
      rows.push([normOrder, item.subbedEquipment, item.qty, item.located, item.quoteReceived, item.runSheet, item.packingSlip, item.notes]);
    });
  });
  const ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  let sheet = ss.getSheetByName(SUB_HELPER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUB_HELPER_SHEET_NAME);
    Logger.log('Scan SUB Sheet: created ' + SUB_HELPER_SHEET_NAME);
  }
  sheet.clear();
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 8).setValues(rows);
  } else {
    sheet.getRange(1, 1, 1, 8).setValues([rows[0]]);
  }
  SpreadsheetApp.flush();
  Logger.log('Scan SUB Sheet: done. ' + orderCount + ' orders, ' + (rows.length - 1) + ' items.');
}
