// --- CONFIGURATION ---
const SHEET_NAME = "CC Connectors"; // The exact name of your inventory sheet tab
const RECIPIENT_EMAILS = [
  "ryan@keslowcamera.com",
  "derek@keslowcamera.com",
  "mikek@keslowcamera.com",
  "chad@keslowcamera.com"
];
const TEST_RECIPIENT = "ryan@keslowcamera.com"; // Only emails Ryan (test report)
const START_ROW = 4; // The first row of data to check (to skip headers)

// Define column numbers for clarity (A=1, B=2, etc.)
const LOCATION_COL = 1;     // Column A
const ITEM_NAME_COL = 2;    // Column B
const PART_NUM_COL = 3;     // Column C
const USAGE_COL = 4;        // Column D
const VENDOR_COL = 6;       // Column F
const THRESHOLD_COL = 9;    // Column I
const QUANTITY_COL = 10;    // Column J
const ORDER_QTY_COL = 11;   // Column K
// ---------------------

/**
 * Scans the inventory sheet for items below their reorder threshold.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The inventory sheet
 * @returns {Array<{location, itemName, partNumber, commonUsage, vendor, orderQty, currentStock}>}
 */
function getLowStockItems(sheet) {
  if (!sheet || sheet.getLastRow() < START_ROW) return [];
  const dataRange = sheet.getRange(START_ROW, 1, sheet.getLastRow(), sheet.getLastColumn());
  const inventoryData = dataRange.getValues();
  const items = [];
  for (const row of inventoryData) {
    const currentSupply = parseFloat(row[QUANTITY_COL - 1]);
    const lowThreshold = parseFloat(row[THRESHOLD_COL - 1]);
    if (!isNaN(currentSupply) && !isNaN(lowThreshold) && currentSupply < lowThreshold) {
      items.push({
        location: row[LOCATION_COL - 1],
        itemName: row[ITEM_NAME_COL - 1],
        partNumber: row[PART_NUM_COL - 1],
        commonUsage: row[USAGE_COL - 1],
        vendor: row[VENDOR_COL - 1],
        orderQty: row[ORDER_QTY_COL - 1],
        currentStock: currentSupply
      });
    }
  }
  return items;
}

/**
 * Builds the HTML table rows for a low-stock email body.
 * @param {Array} itemsToOrder - From getLowStockItems()
 * @returns {string} HTML fragment (table rows only)
 */
function buildLowStockTableRows(itemsToOrder) {
  return itemsToOrder.map(function (item) {
    return `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;">
          <strong>${item.itemName}</strong><br>
          <small><em>Usage: ${item.commonUsage} | Vendor: ${item.vendor}</em></small>
        </td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.partNumber}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.location}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.orderQty}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.currentStock}</td>
      </tr>
    `;
  }).join("");
}

/**
 * Scans the inventory sheet for items below their reorder threshold
 * and sends a single, weekly summary email notification.
 */
function sendWeeklyLowStockReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

    const itemsToOrder = getLowStockItems(sheet);

    if (itemsToOrder.length > 0) {
      const emailSubject = `Weekly Inventory Report: ${itemsToOrder.length} Item(s) Need Reordering`;
      const tableRows = buildLowStockTableRows(itemsToOrder);
      const htmlBody = `
        <p>Hello Team,</p>
        <p>This is your weekly low connector inventory report. The following items are below threshold and it is advised to order more:</p>
        <table style="width:100%; border-collapse: collapse; font-family: sans-serif;">
          <tr style="background-color: #f2f2f2; text-align: left;">
            <th style="padding: 8px; border: 1px solid #ddd;">Item Name</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Part #</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Bin #</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Suggested Re-Order Qty.</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Current Supply</th>
          </tr>
          ${tableRows}
        </table>
        <p>Please review and place the necessary orders. You can view the full inventory sheet <a href="${ss.getUrl()}">here</a>.</p>
        <p>Thank you!</p>
      `;
      GmailApp.sendEmail(RECIPIENT_EMAILS.join(','), emailSubject, "", { htmlBody: htmlBody });
      Logger.log(`Low inventory report sent for ${itemsToOrder.length} items.`);
    } else {
      Logger.log("Weekly inventory check complete. No items are below the reorder threshold.");
    }
  } catch (e) {
    Logger.log(`Error in sendWeeklyLowStockReport function: ${e.message}`);
  }
}

// --- MENU CREATION ---
/**
 * Adds a custom menu to the spreadsheet when it opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Inventory Admin')
    .addItem('Run Test Report (Ryan Only)', 'sendTestInventoryReport')
    .addToUi();
}

/**
 * Scans inventory and sends a TEST report only to the test recipient (Ryan).
 */
function sendTestInventoryReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found.`);
      return;
    }

    const itemsToOrder = getLowStockItems(sheet);

    if (itemsToOrder.length > 0) {
      const emailSubject = `TEST REPORT: Inventory Check - ${itemsToOrder.length} Item(s) Found`;
      const tableRows = buildLowStockTableRows(itemsToOrder);
      const htmlBody = `
        <h3 style="color: red;">*** THIS IS A TEST REPORT ***</h3>
        <p>Hello Ryan,</p>
        <p>This is a manual test of the inventory system. The following items are currently below threshold:</p>
        <table style="width:100%; border-collapse: collapse; font-family: sans-serif;">
          <tr style="background-color: #f2f2f2; text-align: left;">
            <th style="padding: 8px; border: 1px solid #ddd;">Item Name</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Part #</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Bin #</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Rec. Order Qty</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Current Supply</th>
          </tr>
          ${tableRows}
        </table>
        <p>End of Test Report.</p>
      `;
      GmailApp.sendEmail(TEST_RECIPIENT, emailSubject, "", { htmlBody: htmlBody });
      SpreadsheetApp.getUi().alert(`Test Report Sent to ${TEST_RECIPIENT} with ${itemsToOrder.length} items.`);
    } else {
      SpreadsheetApp.getUi().alert("Test complete. No items are currently below the reorder threshold, so no email was sent.");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    Logger.log(e);
  }
}
