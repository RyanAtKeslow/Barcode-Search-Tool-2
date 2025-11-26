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
  
  // ============================================================================
  // Company Holidays Checker (same as in Camera Forecast.js)
  // Company observes: New Year's Day, Memorial Day, Independence Day, Labor Day,
  // Thanksgiving Day, Day after Thanksgiving, Christmas Day, and all days between
  // Christmas and New Year's Day (Christmas week closure)
  // ============================================================================
  function isUSFederalHoliday(date) {
    const year = date.getFullYear();
    const month = date.getMonth(); // 0-based (0 = January)
    const day = date.getDate();
    const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
    
    // Helper to get nth occurrence of a weekday in a month
    function getNthWeekday(year, month, weekday, n) {
      const firstDay = new Date(year, month, 1);
      const firstWeekday = firstDay.getDay();
      let date = 1 + (weekday - firstWeekday + 7) % 7;
      date += (n - 1) * 7;
      return new Date(year, month, date);
    }
    
    // Helper to get last occurrence of a weekday in a month
    function getLastWeekday(year, month, weekday) {
      const lastDay = new Date(year, month + 1, 0); // Last day of month
      const lastWeekday = lastDay.getDay();
      let date = lastDay.getDate() - ((lastWeekday - weekday + 7) % 7);
      return new Date(year, month, date);
    }
    
    // New Year's Day (Jan 1, observed on Friday if Saturday, Monday if Sunday)
    if (month === 0 && day === 1) return true;
    if (month === 0 && day === 2 && dayOfWeek === 1) return true; // Observed Monday if Jan 1 is Sunday
    if (month === 11 && day === 31 && dayOfWeek === 5) return true; // Observed Friday if Jan 1 is Saturday
    
    // Memorial Day (last Monday in May)
    const memorialDay = getLastWeekday(year, 4, 1); // May = month 4
    if (month === memorialDay.getMonth() && day === memorialDay.getDate()) return true;
    
    // Independence Day (July 4, observed on Friday if Saturday, Monday if Sunday)
    if (month === 6 && day === 4) return true;
    if (month === 6 && day === 3 && dayOfWeek === 5) return true; // Observed Friday if July 4 is Saturday
    if (month === 6 && day === 5 && dayOfWeek === 1) return true; // Observed Monday if July 4 is Sunday
    
    // Labor Day (1st Monday in September)
    const laborDay = getNthWeekday(year, 8, 1, 1); // September = month 8
    if (month === laborDay.getMonth() && day === laborDay.getDate()) return true;
    
    // Thanksgiving (4th Thursday in November)
    const thanksgiving = getNthWeekday(year, 10, 4, 4); // November = month 10, Thursday = 4
    if (month === thanksgiving.getMonth() && day === thanksgiving.getDate()) return true;
    
    // Day after Thanksgiving (Friday after Thanksgiving) - Company holiday
    const dayAfterThanksgiving = new Date(thanksgiving);
    dayAfterThanksgiving.setDate(dayAfterThanksgiving.getDate() + 1);
    if (month === dayAfterThanksgiving.getMonth() && day === dayAfterThanksgiving.getDate()) return true;
    
    // Christmas (December 25, observed on Friday if Saturday, Monday if Sunday)
    if (month === 11 && day === 25) return true;
    if (month === 11 && day === 24 && dayOfWeek === 5) return true; // Observed Friday if Dec 25 is Saturday
    if (month === 11 && day === 26 && dayOfWeek === 1) return true; // Observed Monday if Dec 25 is Sunday
    
    // Christmas week closure: All days between Christmas and New Year's Day
    // This includes Dec 26, 27, 28, 29, 30, 31 (and Jan 1 is already handled above)
    if (month === 11 && day >= 26 && day <= 31) return true; // Dec 26-31
    // Note: Jan 1 is already handled above, so we don't need to check it again here
    
    return false;
  }
  
  // Calculate next business day (skip weekends and US Federal Holidays)
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
    
    // Skip holidays (keep advancing until we find a non-holiday, non-weekend day)
    while (isUSFederalHoliday(nextDay) || nextDay.getDay() === 0 || nextDay.getDay() === 6) {
      nextDay.setDate(nextDay.getDate() + 1);
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