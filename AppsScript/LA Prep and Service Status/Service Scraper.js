/**
 * Service Scraper — LA Prep and Service Status
 *
 * Runs on a time-driven trigger (e.g. every minute). Scrapes F2 Imports backup for the most
 * up-to-date service status of assets: either "Serviced for Order" (asset serviced for that
 * specific order) or "Ready to Rent" (RTR) — an asset serviced ahead of time. When a scheduled
 * asset has RTR in F2, that satisfies the requirement for being "Serviced for Order" when
 * that asset is later scheduled for a job.
 *
 * Uses: today's Prep Bay orders, scheduled barcodes per order from Equipment Scheduling Chart
 * (readEquipmentSchedulingData), and F2 Imports backup (Barcode, Order Number, Prep Kind).
 *
 * F2 Imports backup schema (from Process F2 Import): data starts column A.
 * Row 1 = original headers; Row 2 = display; Row 3+ = data.
 * Column C (index 2) = AssetBarcode, G (index 6) = OrderNumber_lu, R (index 17) = PrepKind_lu.
 */

/** This workbook (LA Prep and Service Status) */
const SERVICE_SCRAPER_WORKBOOK_ID = '1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0';
const F2_IMPORTS_BACKUP_SHEET_NAME = 'F2 Imports backup';

/** F2 backup data columns (0-based): Barcode, Order Number, Prep Kind */
const F2_COL_BARCODE = 2;
const F2_COL_ORDER = 6;
const F2_COL_PREP_KIND = 17;

/** Prep Kind values that mean "Ready to Rent" (satisfies serviced-for-order when asset is scheduled) */
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
 * Normalizes a barcode for lookup (lowercase, trim).
 * @param {string} barcode
 * @returns {string}
 */
function normalizeBarcodeForScraper(barcode) {
  return String(barcode || '').trim().toLowerCase();
}

/**
 * Returns true if the Prep Kind value indicates Ready to Rent.
 * @param {string} prepKind
 * @returns {boolean}
 */
function isPrepKindRTR(prepKind) {
  var pk = String(prepKind || '').trim().toLowerCase();
  if (!pk) return false;
  return RTR_PREP_KINDS.some(function (rtr) { return pk === rtr || pk.indexOf(rtr) !== -1; });
}

/**
 * Loads F2 Imports backup and builds a map: barcodeKey -> { orderNumbers: Set, hasRTR: boolean }.
 * One read of columns A–R (minimal range would be C,G,R but we read a block for one round-trip).
 * @returns {Object.<string, { orderNumbers: Set<string>, hasRTR: boolean }>}
 */
function getF2BackupBarcodeStatus() {
  var ss = SpreadsheetApp.openById(SERVICE_SCRAPER_WORKBOOK_ID);
  var sheet = ss.getSheetByName(F2_IMPORTS_BACKUP_SHEET_NAME);
  var out = {};
  if (!sheet) return out;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return out;
  var numRows = lastRow - 2;
  var numCols = 18;
  var values = sheet.getRange(3, 1, numRows, numCols).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
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

/**
 * Returns normalized order numbers that appear in today's Prep Bay (job-block orders).
 * @returns {string[]}
 */
function getPrepBayOrderNumbersForToday() {
  var today = new Date();
  var sheetName = getTodaySheetName(today);
  var prepBayData = readPrepBayDataForDate(sheetName);
  var seen = {};
  var out = [];
  prepBayData.forEach(function (row) {
    var norm = normalizeOrderNumber(row.orderNumber);
    if (norm && !seen[norm]) {
      seen[norm] = true;
      out.push(norm);
    }
  });
  return out;
}

/**
 * For a given order and barcode, checks F2 barcode status. Returns 'serviced_for_order' | 'rtr' | null.
 * @param {Object.<string, { orderNumbers: Set<string>, hasRTR: boolean }>} f2Status
 * @param {string} orderNorm - normalized order number
 * @param {string} barcode - raw barcode (will be normalized for lookup)
 * @returns {string|null} 'serviced_for_order' | 'rtr' | null
 */
function getAssetServiceStatusForOrder(f2Status, orderNorm, barcode) {
  var key = normalizeBarcodeForScraper(barcode);
  var status = f2Status[key];
  if (!status) return null;
  if (status.orderNumbers && status.orderNumbers.has(orderNorm)) return 'serviced_for_order';
  if (status.hasRTR) return 'rtr';
  return null;
}

/**
 * Main entry: scrapes F2 backup and matches scheduled assets (by order) to service status.
 * For each prep bay order and its scheduled barcodes (from Equipment Chart), each asset is
 * either Serviced for Order, satisfied by RTR, or not satisfied.
 * Call from a time-driven trigger (e.g. every minute).
 * @returns {{
 *   byOrder: Object.<string, { servicedForOrder: string[], satisfiedByRTR: string[], notSatisfied: string[] }>,
 *   f2BarcodeCount: number,
 *   prepBayOrderCount: number
 * }}
 */
function runServiceScraper() {
  var f2Status = getF2BackupBarcodeStatus();
  var today = new Date();
  var camerasByOrder = readEquipmentSchedulingData(today);
  var prepBayOrders = getPrepBayOrderNumbersForToday();
  var byOrder = {};
  prepBayOrders.forEach(function (orderNorm) {
    byOrder[orderNorm] = { servicedForOrder: [], satisfiedByRTR: [], notSatisfied: [] };
  });
  Object.keys(camerasByOrder || {}).forEach(function (orderNorm) {
    if (!byOrder[orderNorm]) byOrder[orderNorm] = { servicedForOrder: [], satisfiedByRTR: [], notSatisfied: [] };
    var list = camerasByOrder[orderNorm];
    list.forEach(function (item) {
      var barcode = item.barcode;
      var st = getAssetServiceStatusForOrder(f2Status, orderNorm, barcode);
      if (st === 'serviced_for_order') byOrder[orderNorm].servicedForOrder.push(barcode);
      else if (st === 'rtr') byOrder[orderNorm].satisfiedByRTR.push(barcode);
      else byOrder[orderNorm].notSatisfied.push(barcode);
    });
  });
  var f2BarcodeCount = Object.keys(f2Status).length;
  return {
    byOrder: byOrder,
    f2BarcodeCount: f2BarcodeCount,
    prepBayOrderCount: prepBayOrders.length
  };
}
