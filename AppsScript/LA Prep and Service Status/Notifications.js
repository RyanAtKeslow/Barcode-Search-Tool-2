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
 */
function ensureNotificationsSettingsSheet() {
  var ss = SpreadsheetApp.openById(LA_PREP_STATUS_WORKBOOK_ID);
  var sheet = ss.getSheetByName(NOTIFICATIONS_SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIFICATIONS_SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Orders', 'Serviced Equipment', 'Sub-Rentals']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.getRange(2, 2, 2, 2).setNote('Must be an @keslowcamera.com address. Other addresses are ignored.');
  }
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
  var text = 'New subbed equipment (run sheet confirmed) for today or tomorrow:\n\n';
  newItems.forEach(function (item) {
    text += '• Order ' + item.orderNumber + ' — ' + (item.subbedEquipment || 'Subbed item') + (item.qty ? ' (Qty ' + item.qty + ')' : '') + ' — Prep ' + item.prepDay + '\n';
  });
  if (names.length > 0) text += '\nTracking: ' + names.join(', ');
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
 * Menu: prompt for Google Chat webhook URL and save to Script Properties.
 */
function setNotificationsWebhook() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Google Chat webhook URL', 'Paste the incoming webhook URL for your Chat space. Leave blank to clear.', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var url = result.getResponseText().trim();
  PropertiesService.getScriptProperties().setProperty(NOTIFICATIONS_GCHAT_WEBHOOK_KEY, url);
  ui.alert(url ? 'Webhook URL saved.' : 'Webhook URL cleared.');
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
  template.users = users;
  template.orders = orderList;
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
