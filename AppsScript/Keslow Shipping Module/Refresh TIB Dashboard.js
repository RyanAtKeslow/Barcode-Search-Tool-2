function refreshTIBDashboard() {
  Logger.log('Starting Refresh TIB Dashboard script');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TIB Dashboard');
  if (!sheet) {
    Logger.log('TIB Dashboard sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data to process.');
    return;
  }

  // Get all dates from column G (7), starting from row 2
  const dateRange = sheet.getRange(2, 7, lastRow - 1, 1);
  const dateValues = dateRange.getValues();
  const today = new Date();
  const monthsResults = [];
  const daysResults = [];

  for (let i = 0; i < dateValues.length; i++) {
    const cell = dateValues[i][0];
    let months = '';
    let days = '';
    if (cell && !isNaN(new Date(cell))) {
      const end = new Date(cell);
      months = (end.getFullYear() - today.getFullYear()) * 12 + (end.getMonth() - today.getMonth());
      if (end.getDate() > today.getDate()) months++;
      days = Math.round((end - today) / (1000 * 60 * 60 * 24));
      Logger.log(`Row ${i + 2}: Date ${cell} => ${months} months, ${days} days from today`);
    } else {
      Logger.log(`Row ${i + 2}: Invalid or empty date '${cell}'`);
    }
    monthsResults.push([months]);
    daysResults.push([days]);
  }

  // Write results to column H (8) and I (9), starting from row 2
  sheet.getRange(2, 8, monthsResults.length, 1).setValues(monthsResults);
  sheet.getRange(2, 9, daysResults.length, 1).setValues(daysResults);
  Logger.log('Refresh TIB Dashboard script complete');
} 