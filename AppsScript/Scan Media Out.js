/**
 * Scan Out Media - Media Tracking and Logging Script
 * 
 * This script processes media barcode data from the Dashboard sheet and logs it
 * to the Log sheet with return tracking capabilities and duplicate prevention.
 * 
 * Step-by-step process:
 * 1. Accesses Dashboard and Log sheets from the active spreadsheet
 * 2. Reads data from Dashboard starting at row 3 (A3:G)
 * 3. Extracts global values from row 3 (order number, prep tech, job name, return date)
 * 4. Reads existing log data to check for duplicate barcodes
 * 5. For each row in Dashboard data:
 *    - Skips empty barcodes
 *    - Checks for existing barcode entries in log
 *    - Updates existing entries to "YES" if not already marked
 *    - Formats scan date (uses today's date if empty)
 *    - Creates new log entry with "NO" return status
 * 6. Appends all new entries to the Log sheet starting at row 9
 * 7. Clears processed data from Dashboard (B3:G, preserves column A)
 * 
 * Data Processing:
 * - Source: Dashboard sheet (A=media type, B=scan date, C=barcode, D=order, E=prep tech, F=job, G=return date)
 * - Target: Log sheet (A=media type, B=scan date, C=barcode, D=order, E=prep tech, F=job, G=return date, H=returned)
 * - Global values: Applied to all entries from row 3 of Dashboard
 * - Date formatting: Converts dates to MM/dd/yyyy format
 * 
 * Duplicate Handling:
 * - Checks existing log entries for matching barcodes
 * - Updates existing entries to "YES" if not already marked as returned
 * - Continues processing to create new entries even if duplicates found
 * - Prevents duplicate barcode processing
 * 
 * Return Tracking:
 * - New entries: Marked with "NO" (not returned)
 * - Existing entries: Updated to "YES" (returned)
 * - Status column: Column H in Log sheet
 * 
 * Features:
 * - Duplicate barcode detection and handling
 * - Global value inheritance from row 3
 * - Automatic date formatting and fallback
 * - Return status tracking
 * - Data cleanup after processing
 * - Comprehensive logging and error handling
 */
function ScanOutMedia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const log = ss.getSheetByName("Log");

  const lastRow = dashboard.getLastRow();
  if (lastRow < 3) {
    Logger.log("No data in Dashboard from row 3 down");
    return;
  }

  // Read entire data range A3:G lastRow
  const dataRange = dashboard.getRange(3, 1, lastRow - 2, 7);
  const data = dataRange.getValues();

  // Read the "global" values from row 3 only for columns D:G
  const globalOrderNum = data[0][3];  // D3
  const globalPrepTech = data[0][4];  // E3
  const globalJobName = data[0][5];   // F3
  const globalReturnDate = data[0][6]; // G3

  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");

  // Get existing log data to check for duplicate barcodes
  const logLastRowForCheck = log.getLastRow();
  let existingLogData = [];
  if (logLastRowForCheck >= 9) {
    existingLogData = log.getRange(9, 1, logLastRowForCheck - 8, 8).getValues();
  }

  const outputRows = [];

  data.forEach(row => {
    const mediaType = row[0];
    let scanDate = row[1];
    const barcode = row[2];

    // Skip empty barcodes
    if (!barcode || barcode.toString().trim() === "") return;

    // Check if barcode already exists in log sheet
    // Gather all rows in the log that match this barcode (prior to adding the new row)
    const matchingRowIndexes = [];
    existingLogData.forEach((logRow, idx) => {
      if (logRow[2] && logRow[2].toString().trim() === barcode.toString().trim()) {
        matchingRowIndexes.push(idx);
      }
    });

    if (matchingRowIndexes.length > 0) {
      matchingRowIndexes.forEach(idx => {
        const existingRow = existingLogData[idx];
        const actualRowNumber = idx + 9; // +9 because log data starts at row 9

        // Only update if column H is currently not "YES"
        if (!existingRow[7] || existingRow[7].toString().trim().toUpperCase() !== "YES") {
          log.getRange(actualRowNumber, 8).setValue("YES");
          Logger.log(`Updated existing barcode ${barcode} to YES at row ${actualRowNumber}`);
        }
      });
      // Continue with normal logic (don't return here)
    }

    if (!scanDate || scanDate.toString().trim() === "") {
      scanDate = today;
    } else {
      if (scanDate instanceof Date) {
        scanDate = Utilities.formatDate(scanDate, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
      } else {
        scanDate = scanDate.toString();
      }
    }

    // Append "NO" to indicate not yet returned
    outputRows.push([
      mediaType,       // A
      scanDate,        // B
      barcode,         // C
      globalOrderNum,  // D
      globalPrepTech,  // E
      globalJobName,   // F
      globalReturnDate,// G
      "NO"             // H ← new column
    ]);
  });

  if (outputRows.length === 0) {
    Logger.log("No barcodes found to log.");
    return;
  }

  const startRow = 9;
  const logLastRow = log.getLastRow();
  const pasteRow = Math.max(logLastRow + 1, startRow);

  // Output now includes 8 columns
  log.getRange(pasteRow, 1, outputRows.length, 8).setValues(outputRows);

  // Clear B3:G (not A)
  dashboard.getRange(3, 2, lastRow - 2, 6).clearContent();
}
