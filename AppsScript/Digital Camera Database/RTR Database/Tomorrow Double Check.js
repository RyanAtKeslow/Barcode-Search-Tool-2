/**
 * tomorrowDoubleCheck
 * Scans Prep Bay Assignment sheets to build the list of jobs that start tomorrow
 * and writes them (with header) to columns AJ:AM of the ‘LA Camera Forecast’ sheet.
 */
function tomorrowDoubleCheck() {
  // --- Date helpers ---
  const today = new Date();
  const dayPrefixes = [ 'Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat' ];
  const expectedSheetName = (dateObj) => `${dayPrefixes[dateObj.getDay()]} ${dateObj.getMonth()+1}/${dateObj.getDate()}`;
  
  // Calculate next business day (skip weekends)
  function getNextBusinessDay(date) {
    const nextDay = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayOfWeek = date.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
    
    if (dayOfWeek === 5) { // Friday
      nextDay.setDate(date.getDate() + 3); // Skip to Monday
    } else if (dayOfWeek === 6) { // Saturday
      nextDay.setDate(date.getDate() + 2); // Skip to Monday
    } else if (dayOfWeek === 0) { // Sunday
      nextDay.setDate(date.getDate() + 1); // Skip to Monday
    } else { // Monday-Thursday
      nextDay.setDate(date.getDate() + 1); // Next day
    }
    
    return nextDay;
  }
  
  const nextBusinessDay = getNextBusinessDay(today);

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

  // Locate today & next business day sheets
  const todaySheetName = expectedSheetName(today);
  const nextBusinessDaySheetName = expectedSheetName(nextBusinessDay);
  const todaySheet = visibleSheets.find(sh => sh.getName() === todaySheetName);
  const nextBusinessDaySheet = visibleSheets.find(sh => sh.getName() === nextBusinessDaySheetName);

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
    Logger.log(`tomorrowDoubleCheck: collected ${earlierOrders.size} order numbers & ${earlierJobs.size} job names from today (${todaySheetName}).`);
  } else {
    Logger.log(`tomorrowDoubleCheck: today's sheet "${todaySheetName}" not found.`);
  }

  const outputRows = [];
  const addedEntries = new Set(); // Track complete entries (job+order+camera) to prevent duplicates
  
  if (nextBusinessDaySheet) {
    const rows = nextBusinessDaySheet.getRange('B:J').getValues();
    rows.forEach(r => {
      const jobName    = r[0];
      const orderNum   = normalizeOrder(r[1]);
      const cameraInfo = r[4];
      const note       = r[8];
      if (!jobName && !orderNum) return; // blank

      const lowerName = normalizeString(jobName);

      // Create unique key for this specific entry (job + order + camera info)
      const entryKey = `${lowerName}-${orderNum}-${normalizeString(cameraInfo)}`;

      // Filters
      const wrapOut      = lowerName.includes('wrap out');
      const earlierDup   = earlierOrders.has(orderNum) || earlierJobs.has(lowerName);
      const currentDup   = addedEntries.has(entryKey); // Check for duplicates of this specific entry

      let prepDateTooLate = false;
      if (note && note.toLowerCase().includes('prep')) {
        const matches = note.match(/\d{1,2}\/\d{1,2}/g);
        if (matches) {
          const earliest = matches.reduce((earliest, str) => {
            const [m, d] = str.split('/').map(Number);
            const dt = new Date(nextBusinessDay.getFullYear(), m-1, d);
            // If the date appears to be in the past, it's actually next year
            if (dt < today) {
              dt.setFullYear(nextBusinessDay.getFullYear() + 1);
            }
            return dt < earliest ? dt : earliest;
          }, new Date(nextBusinessDay.getFullYear()+1,0,1));
          // Filter out jobs where prep starts AFTER the next business day
          prepDateTooLate = earliest > nextBusinessDay;
        }
      }

      // Log decision
      Logger.log(`[TomorrowDC] Job:"${jobName}" Ord:${orderNum} Camera:"${cameraInfo}" earlyDup:${earlierDup} currentDup:${currentDup} wrapOut:${wrapOut} prepLate:${prepDateTooLate}`);

      if (!wrapOut && !earlierDup && !currentDup && !prepDateTooLate) {
        outputRows.push([jobName, orderNum, cameraInfo, note]);
        addedEntries.add(entryKey); // Track this specific entry as added
      } else if (currentDup) {
        Logger.log(`[TomorrowDC] Skipping duplicate entry: ${entryKey}`);
      }
    });
  } else {
    Logger.log(`tomorrowDoubleCheck: next business day sheet "${nextBusinessDaySheetName}" not found.`);
  }

  // --- Write to LA Camera Forecast sheet (AJ:AM) ---
  // Note: prepBayBlock() already writes to AI:AM (columns 35-39), so we only need to ensure
  // headers are set in row 1, but we should NOT overwrite the data that prepBayBlock() wrote.
  // Since prepBayBlock() handles the full AI:AM block, tomorrowDoubleCheck() is now redundant
  // for the main output, but we keep it for backward compatibility and any other uses.
  // For now, we'll skip writing to avoid conflicts with prepBayBlock() output.
  
  Logger.log(`tomorrowDoubleCheck: ${outputRows.length} tomorrow jobs found, but output handled by prepBayBlock() to avoid conflicts.`);
  
  // Note: prepBayBlock() already writes tomorrow's jobs as part of its output to AI:AM
  // so we don't need to write separately here to avoid overwriting data.
} 