/**
 * Supply Shelf Labeler
 * @OnlyCurrentDoc
 *
 * Creates a sheet in the target workbook for entering shelf/bin label data and
 * generates a printable 6" x 2.5" label with bin number, name/description,
 * part number, and a QR code (Connector Supplies or General Supplies form).
 */

const SUPPLY_SHELF_LABELER_SHEET_NAME = "Supply Shelf Labeler";
const TARGET_SPREADSHEET_ID = "1BVDeb8sM336-Y5LeQeDd0_p9EA_Sh_HPBCG5Qh5kt_U";

// Column B holds the input values (column A = labels)
const COL_LABELS = 1;
const COL_VALUES = 2;
const ROW_BIN = 1;
const ROW_NAME = 2;
const ROW_PART = 3;
const ROW_LABEL_TYPE = 4;
const ROW_QR_CONNECTOR = 5;
const ROW_QR_GENERAL = 6;

const LABEL_TYPES = ["Connector bin labels", "Shelf labels (General Supplies)"];

/**
 * Creates or gets the Supply Shelf Labeler sheet and sets up headers,
 * data validation for the label-type dropdown, and column widths.
 * Call once when first using the tool (or from menu: Setup sheet).
 */
function setupSupplyShelfLabelerSheet() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SUPPLY_SHELF_LABELER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SUPPLY_SHELF_LABELER_SHEET_NAME);
  }

  // Headers and prompts
  sheet.getRange("A1:B1").setValues([["Shelf / Bin #", ""]]);
  sheet.getRange("A2:B2").setValues([["Name/Description", ""]]);
  sheet.getRange("A3:B3").setValues([["Part Number", ""]]);
  sheet.getRange("A4:B4").setValues([["Label type", ""]]);
  sheet.getRange("A5:B5").setValues([["Connector Supplies form URL", ""]]);
  sheet.getRange("A6:B6").setValues([["General Supplies form URL", ""]]);

  // Dropdown for label type (in B4)
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LABEL_TYPES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("B4").setDataValidation(typeRule);

  const currentType = sheet.getRange("B4").getValue();
  if (!currentType) {
    sheet.getRange("B4").setValue(LABEL_TYPES[0]);
  }

  // Column widths
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 380);

  // Format header column
  sheet.getRange("A1:A6").setFontWeight("bold");

  // Add instructions
  sheet.getRange("A8").setValue("Click the button below to generate a printable label.");
  sheet.getRange("A8").setFontStyle("italic");

  // Add button placeholder note (user assigns script to drawing/button in Sheets UI)
  sheet.getRange("A10").setValue('Assign script: createShelfLabel (Insert → Drawing → assign "createShelfLabel")');
  sheet.getRange("A10").setFontStyle("italic").setFontSize(9);

  SpreadsheetApp.getUi().alert(
    "Sheet ready",
    'Sheet "' + SUPPLY_SHELF_LABELER_SHEET_NAME + '" is set up. Fill in the fields and use the menu "Supply Shelf Labeler" → "Create shelf label" to generate your label.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Supply Shelf Labeler")
    .addItem("Setup sheet", "setupSupplyShelfLabelerSheet")
    .addItem("Create shelf label", "createShelfLabel")
    .addToUi();
}

/**
 * Reads the current sheet values and opens a printable label dialog.
 * Uses the selected label type to choose which QR code (form URL) to show.
 */
function createShelfLabel() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUPPLY_SHELF_LABELER_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      "Sheet not found",
      'Run "Supply Shelf Labeler" → "Setup sheet" first to create the sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const binNumber = String(sheet.getRange(ROW_BIN, COL_VALUES).getValue() || "").trim();
  const nameDesc = String(sheet.getRange(ROW_NAME, COL_VALUES).getValue() || "").trim();
  const partNumber = String(sheet.getRange(ROW_PART, COL_VALUES).getValue() || "").trim();
  const labelType = String(sheet.getRange(ROW_LABEL_TYPE, COL_VALUES).getValue() || "").trim();
  const qrConnectorUrl = String(sheet.getRange(ROW_QR_CONNECTOR, COL_VALUES).getValue() || "").trim();
  const qrGeneralUrl = String(sheet.getRange(ROW_QR_GENERAL, COL_VALUES).getValue() || "").trim();

  if (!binNumber) {
    SpreadsheetApp.getUi().alert("Please enter a Shelf / Bin # in the sheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const isConnector = labelType === LABEL_TYPES[0];
  const qrUrl = isConnector ? qrConnectorUrl : qrGeneralUrl;

  if (!qrUrl) {
    const which = isConnector ? "Connector Supplies" : "General Supplies";
    SpreadsheetApp.getUi().alert(
      "Missing form URL",
      'Please enter the ' + which + ' form URL in the sheet (cell B5 or B6).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const html = buildLabelHtml(binNumber, nameDesc, partNumber, qrUrl);
  const output = HtmlService.createHtmlOutput(html)
    .setWidth(620)
    .setHeight(320)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  SpreadsheetApp.getUi().showModalDialog(output, "Shelf label — Print or save as PDF");
}

/**
 * Builds HTML for a single shelf label: 6" x 2.5", black border, text left, QR right.
 * QR code is loaded from a public API so it works in the dialog.
 */
function buildLabelHtml(binNumber, nameDesc, partNumber, formUrl) {
  const encodedUrl = encodeURIComponent(formUrl);
  const qrImgSrc = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" + encodedUrl;

  // 6in x 2.5in at 96 DPI
  const widthPx = 576;
  const heightPx = 240;

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    'body { font-family: Arial, sans-serif; margin: 0; padding: 16px; background: #eee; }' +
    '.label { width: ' + widthPx + 'px; height: ' + heightPx + 'px; border: 4px solid #000; background: #fff; ' +
    'display: flex; box-sizing: border-box; }' +
    '.label-left { flex: 1; padding: 16px 20px; display: flex; flex-direction: column; justify-content: center; }' +
    '.label-bin { font-size: 36px; font-weight: bold; line-height: 1.2; margin-bottom: 8px; }' +
    '.label-name { font-size: 18px; font-weight: bold; margin-bottom: 6px; }' +
    '.label-part { font-size: 14px; font-weight: bold; }' +
    '.label-right { width: ' + heightPx + 'px; height: ' + heightPx + 'px; flex-shrink: 0; ' +
    'display: flex; align-items: center; justify-content: center; padding: 8px; box-sizing: border-box; }' +
    '.label-right img { max-width: 100%; max-height: 100%; object-fit: contain; }' +
    '.actions { margin-top: 16px; }' +
    '.actions button { padding: 10px 20px; font-size: 14px; cursor: pointer; margin-right: 8px; }' +
    '</style></head><body>' +
    '<div class="label">' +
    '<div class="label-left">' +
    '<div class="label-bin">' + escapeHtml(binNumber) + '</div>' +
    '<div class="label-name">' + escapeHtml(nameDesc) + '</div>' +
    '<div class="label-part">' + escapeHtml(partNumber) + '</div>' +
    '</div>' +
    '<div class="label-right">' +
    '<img src="' + qrImgSrc + '" alt="QR code" />' +
    '</div>' +
    '</div>' +
    '<div class="actions">' +
    '<button onclick="window.print();">Print label</button>' +
    '<button onclick="google.script.host.close();">Close</button>' +
    '</div>' +
    '</body></html>'
  );
}

function escapeHtml(text) {
  const div = { createTextNode: null };
  const el = { innerHTML: "", appendChild: function() {} };
  const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
  return String(text).replace(/[&<>"']/g, function(c) { return map[c] || c; });
}
