/**
 * tomorrowDoubleCheck
 * Scans Prep Bay Assignment sheets to build the list of jobs that start tomorrow
 * and writes them (with header) to columns AJ:AM of the ‘LA Camera Forecast’ sheet.
 */
function tomorrowDoubleCheck() {
  // --- Date helpers ---
  const today     = new Date();
  const tomorrow  = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1); // midnight
  const dayPrefixes = [ 'Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat' ];
  const expectedSheetName = (dateObj) => `${dayPrefixes[dateObj.getDay()]} ${dateObj.getMonth()+1}/${dateObj.getDate()}`;

  // --- Small helpers reused from Camera Forecast script ---
  const normalizeString = (str) => String(str || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const normalizeOrder  = (val) => String(val || '').replace(/[^0-9]/g, '');
  const noteContainsDate = (note, dateObj) => {
    if (!note) return false;
    const matches = note.match(/\d{1,2}\/\d{1,2}/g);
    if (!matches) return false;
    const mTarget = dateObj.getMonth() + 1;
    const dTarget = dateObj.getDate();
    return matches.some(d => {
      const [m, dDay] = d.split('/').map(Number);
      return m === mTarget && dDay === dTarget;
    });
  };

  // --- Open Prep Bay workbook ---
  const prepBaySS = SpreadsheetApp.openById('1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ');
  const visibleSheets = prepBaySS.getSheets().filter(sh => !sh.isSheetHidden());

  // Locate today & tomorrow sheets
  const todaySheetName     = expectedSheetName(today);
  const tomorrowSheetName  = expectedSheetName(tomorrow);
  const todaySheet         = visibleSheets.find(sh => sh.getName() === todaySheetName);
  const tomorrowSheet      = visibleSheets.find(sh => sh.getName() === tomorrowSheetName);

  const earlierOrders = new Set();
  const earlierJobs   = new Set();

  if (todaySheet) {
    const rows = todaySheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName  = r[0];
      const orderNum = normalizeOrder(r[1]);
      if (!jobName && !orderNum) return; // blank row
      earlierOrders.add(orderNum);
      earlierJobs.add(normalizeString(jobName));
    });
    Logger.log(`todayDoubleCheck: collected ${earlierOrders.size} order numbers & ${earlierJobs.size} job names from today (${todaySheetName}).`);
  } else {
    Logger.log(`todayDoubleCheck: today's sheet "${todaySheetName}" not found.`);
  }

  const outputRows = [];
  if (tomorrowSheet) {
    const rows = tomorrowSheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName    = r[0];
      const orderNum   = normalizeOrder(r[1]);
      const cameraInfo = r[4];
      const note       = r[8];
      if (!jobName && !orderNum) return; // blank

      const lowerName = normalizeString(jobName);

      // Filters
      const wrapOut      = lowerName.includes('wrap out');
      const earlierDup   = earlierOrders.has(orderNum) || earlierJobs.has(lowerName);

      let prepDateTooEarly = false;
      if (note && note.toLowerCase().includes('prep')) {
        const matches = note.match(/\d{1,2}\/\d{1,2}/g);
        if (matches) {
          const earliest = matches.reduce((earliest, str) => {
            const [m, d] = str.split('/').map(Number);
            const dt = new Date(tomorrow.getFullYear(), m-1, d);
            return dt < earliest ? dt : earliest;
          }, new Date(tomorrow.getFullYear()+1,0,1));
          prepDateTooEarly = earliest < tomorrow;
        } else {
          prepDateTooEarly = true; // 'prep' mentioned but no date
        }
      }

      // Log decision
      Logger.log(`[TomorrowDC] Job:"${jobName}" Ord:${orderNum} earlyDup:${earlierDup} wrapOut:${wrapOut} prepEarly:${prepDateTooEarly}`);

      if (!wrapOut && !earlierDup && !prepDateTooEarly) {
        outputRows.push([jobName, orderNum, cameraInfo, note]);
      }
    });
  } else {
    Logger.log(`todayDoubleCheck: tomorrow's sheet "${tomorrowSheetName}" not found.`);
  }

  // --- Write to LA Camera Forecast sheet (AJ:AM) ---
  const forecastSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LA Camera Forecast');
  if (!forecastSheet) {
    Logger.log('todayDoubleCheck: LA Camera Forecast sheet not found.');
    return;
  }

  // Clear block
  const maxRows = forecastSheet.getMaxRows();
  forecastSheet.getRange(1, 36, maxRows, 4).clearContent().clearFormat();
  // Header
  forecastSheet.getRange(1, 36, 1, 4).setValues([['Job Name', 'Order #', 'Camera Info', 'Notes']]).setFontWeight('bold');
  // Data
  if (outputRows.length > 0) {
    forecastSheet.getRange(2, 36, outputRows.length, 4).setValues(outputRows);
  }
  Logger.log(`todayDoubleCheck: wrote ${outputRows.length} rows to AJ:AM.`);
} 