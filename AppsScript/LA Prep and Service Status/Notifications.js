/**
 * Notifications — settings sheet, Google Chat webhook, and "pick orders" sidebar.
 * Uses LA_PREP_STATUS_WORKBOOK_ID, SUB_HELPER_SHEET_NAME from the main script.
 * Uses getTodaySheetName, getDateForForecastOffset, readPrepBayDataForDate from the main script.
 */

/** Notifications Settings sheet: Name, Email (@keslowcamera.com only), Orders, Serviced Equipment, Sub-Rentals */
var NOTIFICATIONS_SETTINGS_SHEET_NAME = 'Notifications Settings';
var NOTIFICATIONS_GCHAT_WEBHOOK_KEY = 'NOTIFICATIONS_GCHAT_WEBHOOK';
var KESLOW_EMAIL_SUFFIX = '@keslowcamera.com';

/**
 * Ensures the Notifications Settings sheet exists with headers. Columns: Name, Email, Orders, Serviced Equipment, Sub-Rentals.
 * Column C = order numbers only (no validation). Column B has email note. D and E = checkboxes for opt-in.
 */
function ensureNotificationsSettingsSheet() {
  var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIFICATIONS_SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Orders', 'Serviced Equipment', 'Sub-Rentals']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  // Email note on B2 only (column C is for order numbers, never email validation)
  sheet.getRange(2, 2).setNote('Must be an @keslowcamera.com address. Other addresses are ignored.');
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var normalWidth = 100;
  // Column B: 2.5x width, text clipping
  sheet.setColumnWidth(2, normalWidth * 2.5);
  sheet.getRange(2, 2, Math.max(lastRow, 500), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  // Column C: clear any data validation; set width to 4x normal and text clipping
  var cRange = sheet.getRange(2, 3, lastRow, 1);
  if (cRange) cRange.clearDataValidations();
  sheet.setColumnWidth(3, normalWidth * 4);
  sheet.getRange(2, 3, Math.max(lastRow, 500), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  // Columns D and E only: checkboxes for Serviced Equipment / Sub-Rentals opt-in (getRange 4-arg = row, col, numRows, numColumns)
  var deRange = sheet.getRange(2, 4, Math.max(lastRow, 50), 2);
  deRange.insertCheckboxes();
  // Remove any checkboxes that were wrongly applied to F,G,H in the past
  var fghRange = sheet.getRange(2, 6, Math.max(lastRow, 50), 3);
  fghRange.clearContent().clearDataValidations();
  return sheet;
}

/**
 * Reads Notifications Settings sheet. Returns array of { name, email, orderNumbers, servicedEquipment, subRentals }.
 * Only includes rows where email ends with @keslowcamera.com (blank emails skipped).
 */
function readNotificationSettings() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
    var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
    if (!sheet) return out;
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return out;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var name = row[0] != null ? String(row[0]).trim() : '';
      var email = row[1] != null ? String(row[1]).trim() : '';
      if (!email || !String(email).toLowerCase().endsWith(KESLOW_EMAIL_SUFFIX)) continue;
      var ordersStr = row[2] != null ? String(row[2]).trim() : '';
      var orderNumbers = ordersStr ? ordersStr.split(/[\s,]+/).map(function (s) { return s.replace(/[^0-9]/g, ''); }).filter(function (x) { return !!x; }) : [];
      var servicedEquipment = row[3] === true || row[3] === 'TRUE' || /✓|✔|yes|1/i.test(String(row[3] || ''));
      var subRentals = row[4] === true || row[4] === 'TRUE' || /✓|✔|yes|1/i.test(String(row[4] || ''));
      out.push({ name: name || email, email: email, orderNumbers: orderNumbers, servicedEquipment: servicedEquipment, subRentals: subRentals });
    }
  } catch (e) {
    Logger.log('readNotificationSettings: ' + e.message);
  }
  return out;
}

/**
 * Sends a Google Chat webhook message for new subbed equipment. Finds users who have Sub-Rentals checked and track any of the affected orders.
 */
function sendSubRentalsNotificationToGchat(newItems) {
  if (!newItems || newItems.length === 0) return;
  var webhookUrl = PropertiesService.getScriptProperties().getProperty(NOTIFICATIONS_GCHAT_WEBHOOK_KEY);
  if (!webhookUrl || !webhookUrl.trim()) return;
  var orderNumbers = {};
  newItems.forEach(function (item) { orderNumbers[item.orderNumber] = true; });
  var settings = readNotificationSettings();
  var names = [];
  settings.forEach(function (s) {
    if (!s.subRentals) return;
    var tracking = s.orderNumbers.some(function (ord) { return orderNumbers[ord]; });
    if (tracking) names.push(s.name || s.email);
  });
  var text = 'New subbed equipment (run sheet confirmed) for today:\n\n';
  newItems.forEach(function (item) {
    var jobPart = (item.jobName && String(item.jobName).trim()) ? ' — ' + String(item.jobName).trim() + ' — ' : ' — ';
    text += '• O# ' + item.orderNumber + jobPart + (item.subbedEquipment || 'Subbed item') + (item.qty ? ' (Qty ' + item.qty + ')' : '') + '\n';
  });
  if (names.length > 0) text += '\nTracking: ' + names.map(function (n) { return '@' + n; }).join(', ');
  try {
    UrlFetchApp.fetch(webhookUrl.trim(), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('sendSubRentalsNotificationToGchat: ' + e.message);
  }
}

/**
 * Menu: open sidebar to set Google Chat webhook URL (avoids prompt so execution does not pause).
 */
function setNotificationsWebhook() {
  var html = HtmlService.createHtmlOutputFromFile('WebhookSidebar')
    .setTitle('Set Google Chat webhook')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called from WebhookSidebar: save webhook URL to Script Properties.
 */
function saveNotificationsWebhookUrl(url) {
  var trimmed = (url != null ? String(url) : '').trim();
  PropertiesService.getScriptProperties().setProperty(NOTIFICATIONS_GCHAT_WEBHOOK_KEY, trimmed);
}

/**
 * Returns unique order numbers from Sub Equipment Helper and from Prep Bay for the next 5 workdays (for "pick orders" list).
 * Uses getTodaySheetName, getDateForForecastOffset, readPrepBayDataForDate from main script.
 */
function getCurrentOrderNumbersForTracking() {
  var orders = {};
  try {
    var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
    var helper = ss.getSheetByName(SUB_HELPER_SHEET_NAME);
    if (helper) {
      var data = helper.getDataRange().getValues();
      for (var i = 1; i < (data && data.length); i++) {
        var ord = String(data[i][0] != null ? data[i][0] : '').replace(/[^0-9]/g, '');
        if (ord) orders[ord] = true;
      }
    }
    var today = new Date();
    for (var d = 0; d < 5; d++) {
      var sheetName = getTodaySheetName(getDateForForecastOffset(today, d));
      var prep = readPrepBayDataForDate(sheetName);
      if (prep) prep.forEach(function (r) {
        var ord = String(r.orderNumber || '').replace(/[^0-9]/g, '');
        if (ord) orders[ord] = true;
      });
    }
  } catch (e) {
    Logger.log('getCurrentOrderNumbersForTracking: ' + e.message);
  }
  return Object.keys(orders).sort();
}

/**
 * Menu: open sidebar to pick which orders to track for a selected user.
 */
function showPickOrdersSidebar() {
  ensureNotificationsSettingsSheet();
  var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0] != null ? String(data[i][0]).trim() : '';
    var email = data[i][1] != null ? String(data[i][1]).trim() : '';
    if (name || email) users.push({ index: i + 1, label: (name || email) + (email ? ' (' + email + ')' : '') });
  }
  var orderNumbers = getCurrentOrderNumbersForTracking();
  var orderList = orderNumbers.map(function (o) { return { id: o, label: 'Order ' + o }; });
  var template = HtmlService.createTemplateFromFile('NotificationsSidebar');
  template.sidebarData = { users: users, orders: orderList };
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Pick orders to track').setWidth(320));
}

/**
 * Called from sidebar: save selected order numbers for a user (by 1-based row number).
 */
function saveTrackedOrdersForRow(rowNumber, orderNumbers) {
  var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
  if (!sheet || rowNumber < 2) return;
  var value = Array.isArray(orderNumbers) ? orderNumbers.join(', ') : String(orderNumbers || '').trim();
  sheet.getRange(rowNumber, 3).setValue(value);
}

/**
 * Called from sidebar: get current tracked orders for a user (by 1-based row number).
 */
function getTrackedOrdersForRow(rowNumber) {
  var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
  if (!sheet || rowNumber < 2) return [];
  var value = sheet.getRange(rowNumber, 3).getValue();
  var str = value != null ? String(value).trim() : '';
  return str ? str.split(/[\s,]+/).map(function (s) { return s.replace(/[^0-9]/g, ''); }).filter(function (x) { return !!x; }) : [];
}
