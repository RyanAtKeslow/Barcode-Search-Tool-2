/**
 * Service Scraper — LA Prep and Service Status
 *
 * Builds order numbers from job blocks on all 5 Prep sheets (Prep Today through Prep Four Days Out),
 * finds matching rows in F2 Imports backup (column G = Order Number), and builds a matrix of
 * job-block equipment headers to F2 Equipment Category (column F) for schema follow-up.
 *
 * Uses same logic as Prep Bay Schema Test: getDateForForecastOffset, getTodaySheetName,
 * readPrepBayDataForDate, groupPrepBayByOrder. F2 Imports backup: data from row 3, column F =
 * Equipment Category, column G = Order Number.
 *
 * F2 Imports backup schema: Row 1 = original headers; Row 2 = display headers; Row 3+ = data.
 * Column indices (0-based): C=2 Barcode, F=5 Equipment Category, G=6 Order Number, R=17 Prep Kind.
 */

/** This workbook (LA Prep and Service Status) */
const SERVICE_SCRAPER_WORKBOOK_ID = '1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0';
const F2_IMPORTS_BACKUP_SHEET_NAME = 'F2 Imports backup';

/** F2 backup column indices (0-based). Row 2 display: Prep Date, Service Priority, Barcode, Serial Number, Equipment Name, Equipment Category, Order Number, ... */
const F2_COL_BARCODE = 2;
const F2_COL_EQUIPMENT_NAME = 4;
const F2_COL_EQUIPMENT_CATEGORY = 5;
const F2_COL_ORDER = 6;
const F2_COL_PREP_KIND = 17;

/**
 * Maps F2 Equipment Category (column F) to job-block left-side header. Extend as needed.
 * Cameras are skipped (already filled by Prep Bay Schema Test from Equipment Scheduling Chart).
 */
const F2_EQUIPMENT_CATEGORY_TO_JOB_BLOCK = {
  'Cameras': 'Cameras',
  'Digital Cameras': 'Cameras',
  'Lenses': 'Lenses',
  'Heads': 'Heads',
  'Tripods': 'Heads',
  'Focus': 'Focus',
  'Matte Boxes': 'Matte Boxes',
  'Monitors': 'Monitors',
  'Media': 'Media',
  'Wireless Video': 'Wireless Video',
  'Dir. Viewfinder': 'Dir. Viewfinder',
  'Director Viewfinder': 'Dir. Viewfinder',
  'Ungrouped': 'Ungrouped'
};

/** Prep sheet configs: name and days offset (0 = today, 1 = tomorrow, etc.). Must match PREP_FORECAST_SHEETS in Prep Bay Schema Test. */
const PREP_SHEET_CONFIGS = [
  { name: 'Prep Today', daysOffset: 0 },
  { name: 'Prep Tomorrow', daysOffset: 1 },
  { name: 'Prep Two Days Out', daysOffset: 2 },
  { name: 'Prep Three Days Out', daysOffset: 3 },
  { name: 'Prep Four Days Out', daysOffset: 4 }
];

/** Job block left-side equipment headers (column A labels). Must match EQUIPMENT_CATEGORIES in Prep Bay Schema Test. */
const JOB_BLOCK_EQUIPMENT_HEADERS = [
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

/** Prep Kind values that mean "Ready to Rent" (satisfies serviced-for-order when asset is scheduled). */
const RTR_PREP_KINDS = ['ready to rent', 'rtr'];

/**
 * Normalizes an order number to digits only.
 * @param {string} orderNumber
 * @returns {string}
 */
function normalizeOrderNumber(orderNumber) {
  return String(orderNumber || '').replace(/[^0-9]/g, '').trim();
}

/**
 * Returns order numbers that appear in job blocks for each of the 5 Prep sheets, by sheet name.
 * Uses getDateForForecastOffset, getTodaySheetName, readPrepBayDataForDate, groupPrepBayByOrder (Prep Bay Schema Test).
 * @returns {Object.<string, string[]>} e.g. { 'Prep Today': ['881616', ...], 'Prep Tomorrow': [...], ... }
 */
function getOrderNumbersByDay() {
  var today = new Date();
  var result = {};
  PREP_SHEET_CONFIGS.forEach(function (config) {
    var targetDate = getDateForForecastOffset(today, config.daysOffset);
    var prepBaySheetName = getTodaySheetName(targetDate);
    var prepBayData = readPrepBayDataForDate(prepBaySheetName);
    var jobs = groupPrepBayByOrder(prepBayData || []);
    var orders = [];
    var seen = {};
    jobs.forEach(function (job) {
      var norm = normalizeOrderNumber(job.orderNumber);
      if (norm && !seen[norm]) {
        seen[norm] = true;
        orders.push(norm);
      }
    });
    result[config.name] = orders;
  });
  return result;
}

/**
 * Loads F2 Imports backup and returns all data rows (row 3+) as array of row arrays.
 * Uses getDisplayValues so barcodes and numbers are read as displayed text.
 * @returns {Array.<Array>} Rows of display values; each row is array of column values (0-based indices).
 */
function getF2BackupDataRows() {
  var ss = SpreadsheetApp.openById(SERVICE_SCRAPER_WORKBOOK_ID);
  var sheet = ss.getSheetByName(F2_IMPORTS_BACKUP_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var numCols = 22;
  var values = sheet.getRange(3, 1, lastRow, numCols).getDisplayValues();
  return values;
}

/**
 * Builds set of all order numbers from orderNumbersByDay (all 5 sheets).
 * @param {Object.<string, string[]>} orderNumbersByDay
 * @returns {Object.<string, boolean>} Set-like map of normalized order number -> true
 */
function getAllOrderNumbersSet(orderNumbersByDay) {
  var set = {};
  Object.keys(orderNumbersByDay || {}).forEach(function (sheetName) {
    (orderNumbersByDay[sheetName] || []).forEach(function (norm) {
      if (norm) set[norm] = true;
    });
  });
  return set;
}

/**
 * Maps F2 Equipment Category string to job-block header. Returns 'Ungrouped' if not in map.
 * @param {string} f2Category
 * @returns {string}
 */
function mapF2CategoryToJobBlockHeader(f2Category) {
  var key = String(f2Category || '').trim();
  if (!key) return 'Ungrouped';
  if (F2_EQUIPMENT_CATEGORY_TO_JOB_BLOCK[key]) return F2_EQUIPMENT_CATEGORY_TO_JOB_BLOCK[key];
  return 'Ungrouped';
}

/**
 * Finds F2 backup rows whose Order Number (column G) is in the given set.
 * Commits matching lines to memory with orderNumber, equipmentName, barcode, equipmentCategory, jobBlockHeader.
 * @param {Object.<string, boolean>} orderNumbersSet - normalized order numbers we care about
 * @returns {{ matchingRows: Array, totalRowsRead: number }}
 */
function getF2MatchingRows(orderNumbersSet) {
  var rows = getF2BackupDataRows();
  var matchingRows = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var orderRaw = row[F2_COL_ORDER] != null ? String(row[F2_COL_ORDER]).trim() : '';
    var orderNorm = normalizeOrderNumber(orderRaw);
    if (!orderNumbersSet[orderNorm]) continue;
    var equipmentName = row[F2_COL_EQUIPMENT_NAME] != null ? String(row[F2_COL_EQUIPMENT_NAME]).trim() : '';
    var barcode = row[F2_COL_BARCODE] != null ? String(row[F2_COL_BARCODE]).trim() : '';
    var equipmentCategory = row[F2_COL_EQUIPMENT_CATEGORY] != null ? String(row[F2_COL_EQUIPMENT_CATEGORY]).trim() : '';
    var jobBlockHeader = mapF2CategoryToJobBlockHeader(equipmentCategory);
    matchingRows.push({
      orderNumber: orderNorm,
      equipmentName: equipmentName,
      barcode: barcode,
      equipmentCategory: equipmentCategory,
      jobBlockHeader: jobBlockHeader
    });
  }
  return { matchingRows: matchingRows, totalRowsRead: rows.length };
}

/**
 * Groups F2 matching rows by order and by job-block header. Excludes Cameras (already from Equipment Chart).
 * @param {Array.<{orderNumber: string, equipmentName: string, barcode: string, jobBlockHeader: string}>} matchingRows
 * @returns {Object.<string, Object.<string, Array.<{equipmentName: string, barcode: string}>>>} orderNorm -> header -> [{ equipmentName, barcode }]
 */
function getF2ItemsByOrderByCategory(matchingRows) {
  var byOrder = {};
  (matchingRows || []).forEach(function (r) {
    if (r.jobBlockHeader === 'Cameras') return;
    var norm = r.orderNumber;
    if (!byOrder[norm]) byOrder[norm] = {};
    var header = r.jobBlockHeader;
    if (!byOrder[norm][header]) byOrder[norm][header] = [];
    byOrder[norm][header].push({ equipmentName: r.equipmentName, barcode: r.barcode });
  });
  return byOrder;
}

/**
 * Builds matrix of job-block equipment headers to F2 Equipment Category (column F) values.
 * Returns distinct F2 categories found in matching rows and the list of job block headers;
 * mapping from F2 category -> job block header is left for schema follow-up.
 * @param {Array.<{orderNumber: string, equipmentCategory: string}>} matchingRows
 * @returns {{
 *   jobBlockHeaders: string[],
 *   distinctF2EquipmentCategories: string[],
 *   f2CategoryToJobBlockHeader: Object.<string, string>
 * }}
 */
function buildEquipmentCategoryMatrix(matchingRows) {
  var distinctF2 = {};
  (matchingRows || []).forEach(function (r) {
    var cat = r.equipmentCategory;
    if (cat) distinctF2[cat] = true;
  });
  var distinctF2EquipmentCategories = Object.keys(distinctF2).sort();
  var f2CategoryToJobBlockHeader = {};
  return {
    jobBlockHeaders: JOB_BLOCK_EQUIPMENT_HEADERS.slice(),
    distinctF2EquipmentCategories: distinctF2EquipmentCategories,
    f2CategoryToJobBlockHeader: f2CategoryToJobBlockHeader
  };
}

/** Non-Camera categories (indices 1..9 in JOB_BLOCK_EQUIPMENT_HEADERS). */
var NON_CAMERA_HEADERS = ['Lenses', 'Heads', 'Focus', 'Matte Boxes', 'Monitors', 'Media', 'Wireless Video', 'Dir. Viewfinder', 'Ungrouped'];

/**
 * Parses a Prep sheet to find job blocks (by "End Block" in column A), then for each block finds
 * order number and row index for each equipment category. Returns array of { startRow, endRow, orderNumber, categoryRows }.
 * categoryRows[i] = 1-based row of JOB_BLOCK_EQUIPMENT_HEADERS[i] (Cameras:, Lenses:, ...).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} lastRow
 * @returns {Array.<{startRow: number, endRow: number, orderNumber: string, categoryRows: number[]}>}
 */
function parsePrepSheetBlocks(sheet, lastRow) {
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow, 3).getDisplayValues();
  var endRows = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').indexOf('End Block') !== -1) endRows.push(i + 2);
  }
  var blocks = [];
  for (var b = 0; b < endRows.length; b++) {
    var startRow = b === 0 ? 2 : endRows[b - 1] + 1;
    var endRow = endRows[b];
    var orderNumber = '';
    var categoryRows = [];
    for (var r = startRow - 2; r <= endRow - 2 && r < data.length; r++) {
      var a = String(data[r][0] || '').trim();
      if (a === 'Order #:') orderNumber = normalizeOrderNumber(data[r][1]);
      for (var c = 0; c < JOB_BLOCK_EQUIPMENT_HEADERS.length; c++) {
        if (a === JOB_BLOCK_EQUIPMENT_HEADERS[c] + ':') categoryRows[c] = r + 2;
      }
    }
    blocks.push({ startRow: startRow, endRow: endRow, orderNumber: orderNumber, categoryRows: categoryRows });
  }
  return blocks;
}

/**
 * Writes F2 serviced equipment (Equipment Name, Barcode) into one Prep sheet. For each job block,
 * fills columns B and C for non-Camera categories; when a category has more than one item, inserts
 * a row after the first and moves everything below down, then writes the extra item(s).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} sheetName - for logging
 * @param {string[]} orderNumbersOnSheet - normalized order numbers that appear on this sheet
 * @param {Object.<string, Object.<string, Array>>} f2ByOrderByCat - from getF2ItemsByOrderByCategory
 */
function writeF2ServicedEquipmentToPrepSheet(sheet, sheetName, orderNumbersOnSheet, f2ByOrderByCat) {
  var orderSet = {};
  orderNumbersOnSheet.forEach(function (n) { orderSet[n] = true; });
  var lastRow = Math.min(sheet.getLastRow(), 2500);
  var blocks = parsePrepSheetBlocks(sheet, lastRow);
  Logger.log('  [' + sheetName + '] Parsed ' + blocks.length + ' job block(s), ' + orderNumbersOnSheet.length + ' order(s) on sheet');
  var blocksUpdated = 0;
  var totalItemsWritten = 0;
  var totalRowsInserted = 0;
  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    var orderNorm = block.orderNumber;
    if (!orderNorm || !orderSet[orderNorm]) continue;
    var itemsByCat = f2ByOrderByCat[orderNorm];
    if (!itemsByCat) continue;
    blocksUpdated++;
    var blockItems = 0;
    var blockInserted = 0;
    for (var h = 0; h < NON_CAMERA_HEADERS.length; h++) {
      var header = NON_CAMERA_HEADERS[h];
      var catIndex = JOB_BLOCK_EQUIPMENT_HEADERS.indexOf(header);
      if (catIndex < 0) continue;
      var row = block.categoryRows[catIndex];
      if (!row) continue;
      var items = itemsByCat[header];
      if (!items || items.length === 0) continue;
      var numCols = 10;
      var jobBg = sheet.getRange(row, 1).getBackground();
      sheet.getRange(row, 2).setValue(items[0].equipmentName);
      sheet.getRange(row, 3).setValue(items[0].barcode);
      blockItems += 1;
      for (var i = 1; i < items.length; i++) {
        sheet.insertRowAfter(row);
        sheet.getRange(row + 1, 2).setValue(items[i].equipmentName);
        sheet.getRange(row + 1, 3).setValue(items[i].barcode);
        sheet.getRange(row + 1, 1, 1, numCols).setBackground(jobBg);
        row += 1;
        blockItems++;
        blockInserted++;
        for (var j = catIndex + 1; j < block.categoryRows.length; j++) {
          if (block.categoryRows[j] != null) block.categoryRows[j]++;
        }
        block.endRow++;
        for (var k = b + 1; k < blocks.length; k++) {
          blocks[k].startRow++;
          blocks[k].endRow++;
          for (var q = 0; q < blocks[k].categoryRows.length; q++) {
            if (blocks[k].categoryRows[q] != null) blocks[k].categoryRows[q]++;
          }
        }
      }
    }
    totalItemsWritten += blockItems;
    totalRowsInserted += blockInserted;
    if (blockItems > 0) {
      Logger.log('    Order ' + orderNorm + ': ' + blockItems + ' item(s) written' + (blockInserted > 0 ? ', ' + blockInserted + ' row(s) inserted' : ''));
    }
  }
  Logger.log('  [' + sheetName + '] Updated ' + blocksUpdated + ' block(s), ' + totalItemsWritten + ' item(s) written, ' + totalRowsInserted + ' row(s) inserted');
}

/**
 * Deletes all rows beyond row 2000 on the given sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number} Number of rows deleted (0 if none)
 */
function deleteRowsBeyond2000(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2000) return 0;
  var maxRows = sheet.getMaxRows();
  var numToDelete = maxRows - 2000;
  sheet.deleteRows(2001, numToDelete);
  return numToDelete;
}

/**
 * Deletes rows beyond row 2000 for all 5 Prep forecast sheets.
 */
function deleteAllPrepSheetsRowsBeyond2000() {
  var ss = SpreadsheetApp.openById(SERVICE_SCRAPER_WORKBOOK_ID);
  Logger.log('Trimming Prep sheets to 2000 rows max:');
  PREP_SHEET_CONFIGS.forEach(function (config) {
    var sheet = ss.getSheetByName(config.name);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow <= 2000) {
        Logger.log('  ' + config.name + ': ' + lastRow + ' rows (no trim)');
      } else {
        var deleted = deleteRowsBeyond2000(sheet);
        Logger.log('  ' + config.name + ': trimmed to 2000 rows (deleted ' + deleted + ')');
      }
    }
  });
}

/**
 * Main entry: build order numbers by day, F2 matching rows, write Equipment Name/Barcode into
 * Serviced Equipment area (B/C) of each job block on all 5 Prep sheets (excluding Cameras),
 * inserting rows when a category has multiple items; then trim all Prep sheets to 2000 rows.
 * Call from menu or time-driven trigger.
 * @returns {{
 *   orderNumbersByDay: Object.<string, string[]>,
 *   allOrderNumbersSet: Object.<string, boolean>,
 *   f2MatchingRows: Array,
 *   equipmentCategoryMatrix: Object,
 *   writtenSheets: string[]
 * }}
 */
function runServiceScraper() {
  Logger.log('--- Service Scraper started ---');

  var orderNumbersByDay = getOrderNumbersByDay();
  Logger.log('Order numbers by day:');
  Object.keys(orderNumbersByDay).forEach(function (name) {
    Logger.log('  ' + name + ': ' + (orderNumbersByDay[name].length) + ' order(s)');
  });

  var allOrderNumbersSet = getAllOrderNumbersSet(orderNumbersByDay);
  var totalOrders = Object.keys(allOrderNumbersSet).length;
  Logger.log('Total unique orders (all 5 sheets): ' + totalOrders);

  var matchResult = getF2MatchingRows(allOrderNumbersSet);
  var matchingRows = matchResult.matchingRows;
  Logger.log('F2 backup: ' + (matchResult.totalRowsRead || 0) + ' data row(s) read, ' + matchingRows.length + ' matched to prep orders');

  var equipmentCategoryMatrix = buildEquipmentCategoryMatrix(matchingRows);
  if (equipmentCategoryMatrix.distinctF2EquipmentCategories.length > 0) {
    Logger.log('F2 Equipment Categories in matches: ' + equipmentCategoryMatrix.distinctF2EquipmentCategories.join(', '));
  }

  var f2ByOrderByCat = getF2ItemsByOrderByCategory(matchingRows);
  var ordersWithF2Data = Object.keys(f2ByOrderByCat).length;
  Logger.log('Orders with F2 serviced equipment (non-Camera): ' + ordersWithF2Data);

  Logger.log('Writing to Prep sheets:');
  var ss = SpreadsheetApp.openById(SERVICE_SCRAPER_WORKBOOK_ID);
  var writtenSheets = [];
  PREP_SHEET_CONFIGS.forEach(function (config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      Logger.log('  [' + config.name + '] Sheet not found, skipped');
      return;
    }
    var ordersOnSheet = orderNumbersByDay[config.name] || [];
    writeF2ServicedEquipmentToPrepSheet(sheet, config.name, ordersOnSheet, f2ByOrderByCat);
    writtenSheets.push(config.name);
  });

  deleteAllPrepSheetsRowsBeyond2000();

  Logger.log('--- Service Scraper completed ---');

  return {
    orderNumbersByDay: orderNumbersByDay,
    allOrderNumbersSet: allOrderNumbersSet,
    f2MatchingRows: matchingRows,
    equipmentCategoryMatrix: equipmentCategoryMatrix,
    writtenSheets: writtenSheets
  };
}

// --- Legacy helpers (kept for compatibility if anything still uses barcode/order/RTR status) ---

function normalizeBarcodeForScraper(barcode) {
  return String(barcode || '').trim().toLowerCase();
}

function isPrepKindRTR(prepKind) {
  var pk = String(prepKind || '').trim().toLowerCase();
  if (!pk) return false;
  return RTR_PREP_KINDS.some(function (rtr) { return pk === rtr || pk.indexOf(rtr) !== -1; });
}

/** F2 backup: barcodeKey -> { orderNumbers: Set, hasRTR: boolean }. Uses same column indices. */
function getF2BackupBarcodeStatus() {
  var rows = getF2BackupDataRows();
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var barcodeRaw = row[F2_COL_BARCODE] != null ? String(row[F2_COL_BARCODE]).trim() : '';
    var barcodeKey = normalizeBarcodeForScraper(barcodeRaw);
    if (!barcodeKey) continue;
    if (!out[barcodeKey]) out[barcodeKey] = { orderNumbers: new Set(), hasRTR: false };
    var orderNorm = normalizeOrderNumber(row[F2_COL_ORDER]);
    if (orderNorm) out[barcodeKey].orderNumbers.add(orderNorm);
    if (isPrepKindRTR(row[F2_COL_PREP_KIND])) out[barcodeKey].hasRTR = true;
  }
  return out;
}

function getAssetServiceStatusForOrder(f2Status, orderNorm, barcode) {
  var key = normalizeBarcodeForScraper(barcode);
  var status = f2Status[key];
  if (!status) return null;
  if (status.orderNumbers && status.orderNumbers.has(orderNorm)) return 'serviced_for_order';
  if (status.hasRTR) return 'rtr';
  return null;
}
