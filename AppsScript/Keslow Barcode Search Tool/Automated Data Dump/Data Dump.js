function F2DataDumpDirectPrint() {
  let clearedData = {};
  const sheetsToProcess = ['Barcode SEARCH', 'Receiving Bays', 'Prep Bays'];
  const devMode = false; // Set to true for development mode
  let processedThread = null; // Track the processed thread for marking as read
  let foundMatchingEmail = false; // Track if we found a matching email

  const restoreClearedData = (ss, clearedData) => {
    sheetsToProcess.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || !clearedData[sheetName]) return;
      
      // Create a map of cells that had data
      const cellMap = new Map();
      
      // Fill in the data from clearedData
      clearedData[sheetName].forEach(({ col, startRow, values }) => {
        if (values.length === 0) return;
        values.forEach((value, i) => {
          if (value[0] !== '') {  // Only store non-empty values
            const row = startRow + i; // Fix: remove the -1 since we want to start from startRow
            const key = `${row},${col}`;
            cellMap.set(key, value[0]);
          }
        });
      });
      
      if (cellMap.size === 0) {
        Logger.log(`ℹ️ No data to restore for ${sheetName}`);
        return;
      }
      
      // Convert map to ranges for batch writing
      const ranges = [];
      let currentRange = null;
      
      // Sort keys to process rows and columns in order
      const sortedKeys = Array.from(cellMap.keys()).sort((a, b) => {
        const [rowA, colA] = a.split(',').map(Number);
        const [rowB, colB] = b.split(',').map(Number);
        return rowA === rowB ? colA - colB : rowA - rowB;
      });
      
      sortedKeys.forEach(key => {
        const [row, col] = key.split(',').map(Number);
        const value = cellMap.get(key);
        
        if (!currentRange) {
          currentRange = {
            startRow: row,
            startCol: col,
            endRow: row,
            endCol: col,
            values: [[value]]
          };
        } else if (row === currentRange.endRow && col === currentRange.endCol + 1) {
          // Extend current range horizontally
          currentRange.endCol = col;
          currentRange.values[0].push(value);
        } else {
          // Save current range and start new one
          ranges.push(currentRange);
          currentRange = {
            startRow: row,
            startCol: col,
            endRow: row,
            endCol: col,
            values: [[value]]
          };
        }
      });
      
      if (currentRange) {
        ranges.push(currentRange);
      }
      
      // Write each range
      ranges.forEach(range => {
        const numRows = range.endRow - range.startRow + 1;
        const numCols = range.endCol - range.startCol + 1;
        sheet.getRange(range.startRow, range.startCol, numRows, numCols)
             .setValues(range.values);
      });
      
      Logger.log(`🔄 Restored ${cellMap.size} cells in ${ranges.length} ranges for ${sheetName}`);
    });
    Logger.log("✅ All cleared bay data restored.");
  };

  let summaryStats = {
    totalRows: 0,
    barcodeCount: 0,
    headerRow: '',
    processedBarcodes: '',
    sourceRows: 0,
    sourceBarcodeCount: 0
  };

  try {
    const writeChunkSize = 10000;
    const maxRetries = 3;
    const retryDelay = 5000;
    const recipient = "Share@keslowcamera.com";

    Logger.log("🔍 Getting active spreadsheet...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Could not access active spreadsheet");
    }
    const targetSheetName = 'Barcode Dictionary';
    Logger.log("🔍 About to get sheet by name...");
    let barcodeSheet = ss.getSheetByName(targetSheetName);
    Logger.log("🔍 getSheetByName finished.");
    if (!barcodeSheet) {
      Logger.log("🔍 Sheet not found, about to insert sheet...");
      barcodeSheet = ss.insertSheet(targetSheetName);
      Logger.log("🔍 insertSheet finished.");
    }
    Logger.log("✅ Target sheet ready.");

    let allData;
    if (devMode) {
      Logger.log("🛠️ Dev mode enabled: Starting bay clearing process...");
      
      // Define getBayRanges function at the top of dev mode scope
      const getBayRanges = (sheetName, sheet) => {
        // Returns an array of {rangeA1, col, startRow} objects for each bay in the sheet
        const ranges = [];
        const lastRow = sheet.getLastRow();
        if (sheetName === 'Barcode SEARCH' || sheetName === 'Receiving Bays') {
          // Get the last column of the sheet
          const lastCol = sheet.getLastColumn();
          
          // Calculate columns to clear (every 4th column starting from A)
          for (let col = 1; col <= lastCol; col += 4) {
            const colLetter = String.fromCharCode(64 + col);
            ranges.push({ rangeA1: `${colLetter}4:${colLetter}`, col, startRow: 4 });
          }
        } else if (sheetName === 'Prep Bays') {
          // Get the last column of the sheet
          const lastCol = sheet.getLastColumn();
          
          // Calculate columns to clear (every 3rd column starting from A)
          for (let col = 1; col <= lastCol; col += 3) {
            const colLetter = String.fromCharCode(64 + col);
            ranges.push({ rangeA1: `${colLetter}4:${colLetter}`, col, startRow: 4 });
          }
        }
        return ranges;
      };

      // --- Step 1: Clear and store bay ranges BEFORE any other work ---
      Logger.log("🧹 Clearing and storing bay data...");
      clearedData = {};
      sheetsToProcess.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        const ranges = getBayRanges(sheetName, sheet);
        clearedData[sheetName] = [];
        ranges.forEach(({ rangeA1, col, startRow }) => {
          const lastRow = sheet.getLastRow();
          if (lastRow < startRow) return;
          const numRows = lastRow - startRow + 1;
          const range = sheet.getRange(startRow, col, numRows, 1);
          const values = range.getValues();
          clearedData[sheetName].push({ col, startRow, values });
          range.clearContent();
          Logger.log(`🧹 Cleared ${sheetName} col ${col} from row ${startRow} (${numRows} rows)`);
        });
      });
      Logger.log("✅ Bay data cleared and stored in memory.");

      Logger.log("🛠️ Dev mode enabled: Skipping email fetch and conversion. Fetching most recent 'Assets_GoogleExport' file from Drive...");
      let latestFile = null;
      let latestDate = 0;
      const search = DriveApp.searchFiles("title contains 'Assets_GoogleExport'");
      while (search.hasNext()) {
        const file = search.next();
        if (file.getLastUpdated().getTime() > latestDate) {
          latestFile = file;
          latestDate = file.getLastUpdated().getTime();
        }
      }
      if (!latestFile) {
        Logger.log("❌ No file found with 'Assets_GoogleExport' in title.");
        return;
      }
      Logger.log(`🛠️ Found file: ${latestFile.getName()} (ID: ${latestFile.getId()})`);
      Logger.log(`📄 Exact filename: ${latestFile.getName()}`);
      let sheetFileId;
      if (latestFile.getMimeType() === MimeType.GOOGLE_SHEETS) {
        Logger.log("🛠️ File is already a Google Sheet. Using directly.");
        sheetFileId = latestFile.getId();
      } else {
        Logger.log("🛠️ File is not a Google Sheet. Converting...");
        const convertedFile = Drive.Files.copy({
          title: latestFile.getName(),
          mimeType: MimeType.GOOGLE_SHEETS
        }, latestFile.getId());
        Logger.log("🛠️ Converted to Google Sheet: " + convertedFile.id);
        sheetFileId = convertedFile.id;
      }
      const convertedSheet = SpreadsheetApp.openById(sheetFileId);
      const sourceSheet = convertedSheet.getSheets()[0];
      allData = sourceSheet.getDataRange().getValues();
      Logger.log(`📊 Source sheet has ${sourceSheet.getLastRow()} rows`);
      Logger.log("📊 Reading source data...");
      const sourceLastRow = sourceSheet.getLastRow();
      const sourceBarcodeCount = sourceSheet.getRange("C2:C" + sourceLastRow).getValues().flat().filter(cell => cell).length;
      Logger.log(`📊 Source sheet statistics:\n- Total rows: ${sourceLastRow}\n- Data rows: ${sourceLastRow - 1}\n- Raw barcode count: ${sourceBarcodeCount}`);
      summaryStats.sourceRows = sourceLastRow;
      summaryStats.sourceBarcodeCount = sourceBarcodeCount;
      Logger.log("🔄 Processing and formatting data...");
      const processedData = sortFlawlessDataAutomationMode(allData);
      if (!processedData || processedData.length === 0) {
        Logger.log("⚠️ No data returned from sortFlawlessDataAutomationMode.");
        return;
      }
      Logger.log(`📊 Processed data statistics:\n- Total rows: ${processedData.length}\n- Header row: ${processedData[0].join(', ')}`);
      Logger.log("🔍 Verifying processed data...");
      Logger.log(`Processed data: ${processedData.map(row => row[6]).join(', ')}`);
      summaryStats.totalRows = processedData.length;
      summaryStats.headerRow = processedData[0].join(', ');
      summaryStats.processedBarcodes = processedData.map(row => row[6]).join(', ');
      summaryStats.barcodeCount = processedData.map(row => row[6]).filter(Boolean).length;
      Logger.log("🧹 Clearing Barcode Dictionary sheet...");
      barcodeSheet.clearContents();
      barcodeSheet.clearFormats();
      Logger.log("📝 Writing processed data directly to Barcode Dictionary...");
      // Insert completion message in A1, headers to row 2, data starts from row 4
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM');
      barcodeSheet.clearContents();
      barcodeSheet.clearFormats();
      barcodeSheet.getRange(1, 1).setValue(`Data Export Completed on ${today}`);
      barcodeSheet.getRange(2, 1, processedData.length, processedData[0].length).setValues(processedData);
      barcodeSheet.setFrozenRows(2);
      Logger.log("✅ Processed data written directly to Barcode Dictionary with completion message and frozen rows.");
      // Calculate processed barcode count (pipes + 1 in each cell in column G, skipping header)
      let processedBarcodeCount = 0;
      const dataRows = barcodeSheet.getRange(4, 7, processedData.length - 2, 1).getValues(); // Column G, skip header, start from row 4
      dataRows.forEach(row => {
        const cell = row[0] ? row[0].toString() : '';
        if (cell.length > 0) {
          processedBarcodeCount += cell.split('|').length;
        }
      });
      summaryStats.barcodeCount = processedBarcodeCount;
      
      // --- Secondary Database Export ---
      Logger.log("🔄 Starting secondary database export...");
      const secondaryExportResult = exportToSecondaryDatabase(allData, summaryStats);
      if (secondaryExportResult.success) {
        Logger.log(`✅ Secondary export completed: ${secondaryExportResult.rowsExported} rows exported to ${secondaryExportResult.targetSpreadsheet}`);
        summaryStats.secondaryExport = secondaryExportResult;
      } else {
        Logger.log(`❌ Secondary export failed: ${secondaryExportResult.message}`);
        summaryStats.secondaryExport = secondaryExportResult;
      }
      
      return;
    } else {
      Logger.log("📧 Searching for unread emails...");
      const threads = GmailApp.search('is:unread subject:"Assets Excel Export for Google"');
      if (!threads.length) {
        Logger.log("📭 No matching unread email found. Exiting quietly.");
        return;
      }
      foundMatchingEmail = true;
      Logger.log(`📧 Found ${threads.length} matching email threads`);
      Logger.log("📡 Barcode automation started.");
      MailApp.sendEmail({
        to: recipient,
        subject: "📡 Barcode Automation Started",
        body: `The barcode automation started running at ${new Date().toLocaleString()}.`
      });

      // Define getBayRanges function for non-dev mode
      const getBayRanges = (sheetName, sheet) => {
        // Returns an array of {rangeA1, col, startRow} objects for each bay in the sheet
        const ranges = [];
        const lastRow = sheet.getLastRow();
        if (sheetName === 'Barcode SEARCH' || sheetName === 'Receiving Bays') {
          // Get the last column of the sheet
          const lastCol = sheet.getLastColumn();
          
          // Calculate columns to clear (every 4th column starting from A)
          for (let col = 1; col <= lastCol; col += 4) {
            const colLetter = String.fromCharCode(64 + col);
            ranges.push({ rangeA1: `${colLetter}4:${colLetter}`, col, startRow: 4 });
          }
        } else if (sheetName === 'Prep Bays') {
          // Get the last column of the sheet
          const lastCol = sheet.getLastColumn();
          
          // Calculate columns to clear (every 3rd column starting from A)
          for (let col = 1; col <= lastCol; col += 3) {
            const colLetter = String.fromCharCode(64 + col);
            ranges.push({ rangeA1: `${colLetter}4:${colLetter}`, col, startRow: 4 });
          }
        }
        return ranges;
      };

      // --- Step 1: Clear and store bay ranges BEFORE any other work ---
      Logger.log("🧹 Clearing and storing bay data...");
      clearedData = {};
      sheetsToProcess.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        const ranges = getBayRanges(sheetName, sheet);
        clearedData[sheetName] = [];
        ranges.forEach(({ rangeA1, col, startRow }) => {
          const lastRow = sheet.getLastRow();
          if (lastRow < startRow) return;
          const numRows = lastRow - startRow + 1;
          const range = sheet.getRange(startRow, col, numRows, 1);
          const values = range.getValues();
          clearedData[sheetName].push({ col, startRow, values });
          range.clearContent();
          Logger.log(`🧹 Cleared ${sheetName} col ${col} from row ${startRow} (${numRows} rows)`);
        });
      });
      Logger.log("✅ Bay data cleared and stored in memory.");

      Logger.log("🚀 Starting F2DataDumpDirectPrint function...");
      Logger.log("📧 Processing email threads...");
      const startTime = new Date().getTime();
      const MAX_RUNTIME_MS = 4 * 60 * 1000; // 4 minutes
      for (const thread of threads) {
        Logger.log("📧 Getting messages from thread...");
        const messages = thread.getMessages();
        Logger.log(`📧 Found ${messages.length} messages in thread`);
        if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
          Logger.log('⏳ Approaching timeout, exiting early to avoid session hang.');
          return;
        }
        for (const message of messages) {
          Logger.log("📧 Checking message attachments...");
          const attachments = message.getAttachments();
          Logger.log(`📧 Found ${attachments.length} attachments`);
          const attachment = attachments.find(att => {
            const name = att.getName().toLowerCase();
            Logger.log(`📎 Checking attachment: ${name}`);
            return name.endsWith('.xlsx');
          });
          if (!attachment) {
            Logger.log("📎 No Excel attachment found in this message");
            continue;
          }
          try {
            // Upload attachment to Drive
            const uploadedFile = DriveApp.createFile(attachment.copyBlob());
            Logger.log("📁 Uploaded file to Drive: " + uploadedFile.getName());
            Logger.log(`📊 File size: ${uploadedFile.getSize()} bytes`);
            // Initial wait before conversion for large files
            if (uploadedFile.getSize() > 1000000) { // If file is larger than 1MB
              Logger.log("⏳ Large file detected, waiting 10 seconds before conversion...");
              Utilities.sleep(10000);
            }
            // Convert uploaded file to Google Sheets
            const convertedFile = Drive.Files.copy({
              title: uploadedFile.getName(),
              mimeType: MimeType.GOOGLE_SHEETS
            }, uploadedFile.getId());
            Logger.log("🔄 Converted to Google Sheet: " + convertedFile.id);
            // Wait until the converted file is fully ready
            let ready = false, attempts = 0, maxAttempts = 30;
            const waitTime = 10000;
            // Initial wait after conversion
            Logger.log("⏳ Waiting 20 seconds for initial conversion processing...");
            Utilities.sleep(20000);
            while (!ready && attempts++ < maxAttempts) {
              try {
                const fileSize = DriveApp.getFileById(convertedFile.id).getSize();
                Logger.log(`⏳ Attempt ${attempts}: File size is ${fileSize} bytes`);
                // Try to open the sheet to verify it's really ready
                try {
                  const testSheet = SpreadsheetApp.openById(convertedFile.id);
                  const testRange = testSheet.getSheets()[0].getRange("A1").getValue();
                  ready = true;
                  Logger.log("✅ File is ready and accessible.");
                } catch (e) {
                  Logger.log(`⚠️ File not yet accessible: ${e.toString()}`);
                  ready = false;
                }
              } catch (e) {
                Logger.log(`⚠️ Attempt ${attempts} failed to get file size: ${e.toString()}`);
              }
              if (!ready) {
                Logger.log(`⏳ Waiting ${waitTime/1000} seconds before next attempt...`);
                Utilities.sleep(waitTime);
              }
            }
            if (!ready) {
              throw new Error(`❌ Conversion timeout: File not ready after ${(maxAttempts * waitTime)/1000} seconds.`);
            }
            Logger.log("🔍 Opening converted sheet...");
            const convertedSheet = SpreadsheetApp.openById(convertedFile.id);
            const sourceSheet = convertedSheet.getSheets()[0];
            Logger.log(`📊 Source sheet has ${sourceSheet.getLastRow()} rows`);
            // Get all data at once
            Logger.log("📊 Reading source data...");
            allData = sourceSheet.getDataRange().getValues();
            // Add source sheet statistics
            const sourceLastRow = sourceSheet.getLastRow();
            const sourceBarcodeCount = sourceSheet.getRange("C2:C" + sourceLastRow).getValues().flat().filter(cell => cell).length;
            Logger.log(`📊 Source sheet statistics:\n- Total rows: ${sourceLastRow}\n- Data rows: ${sourceLastRow - 1}\n- Raw barcode count: ${sourceBarcodeCount}`);
            summaryStats.sourceRows = sourceLastRow;
            summaryStats.sourceBarcodeCount = sourceBarcodeCount;
            processedThread = thread; // Track the processed thread
            break;
          } catch (error) {
            Logger.log(`❌ Error processing file: ${error.toString()}`);
            Logger.log(`Stack trace: ${error.stack}`);
            throw error;
          }
        }
        if (allData) break;
      }
      // Continue with the rest of the script using allData...
      Logger.log("🔄 Processing and formatting data...");
      const processedData = sortFlawlessDataAutomationMode(allData);
      if (!processedData || processedData.length === 0) {
        Logger.log("⚠️ No data returned from sortFlawlessDataAutomationMode.");
        return;
      }
      Logger.log(`📊 Processed data statistics:\n- Total rows: ${processedData.length}\n- Header row: ${processedData[0].join(', ')}`);
      Logger.log("🔍 Verifying processed data...");
      Logger.log(`Processed data: ${processedData.map(row => row[6]).join(', ')}`);
      summaryStats.totalRows = processedData.length;
      summaryStats.headerRow = processedData[0].join(', ');
      summaryStats.processedBarcodes = processedData.map(row => row[6]).join(', ');
      summaryStats.barcodeCount = processedData.map(row => row[6]).filter(Boolean).length;
      Logger.log("🧹 Clearing Barcode Dictionary sheet...");
      barcodeSheet.clearContents();
      barcodeSheet.clearFormats();
      Logger.log("📝 Writing processed data directly to Barcode Dictionary...");
      // Insert completion message in A1, headers to row 2, data starts from row 4
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM');
      barcodeSheet.clearContents();
      barcodeSheet.clearFormats();
      barcodeSheet.getRange(1, 1).setValue(`Data Export Completed on ${today}`);
      barcodeSheet.getRange(2, 1, processedData.length, processedData[0].length).setValues(processedData);
      barcodeSheet.setFrozenRows(2);
      Logger.log("✅ Processed data written directly to Barcode Dictionary with completion message and frozen rows.");
      // Calculate processed barcode count (pipes + 1 in each cell in column G, skipping header)
      let processedBarcodeCount = 0;
      const dataRows = barcodeSheet.getRange(4, 7, processedData.length - 2, 1).getValues(); // Column G, skip header, start from row 4
      dataRows.forEach(row => {
        const cell = row[0] ? row[0].toString() : '';
        if (cell.length > 0) {
          processedBarcodeCount += cell.split('|').length;
        }
      });
      summaryStats.barcodeCount = processedBarcodeCount;
      
      // --- Secondary Database Export ---
      Logger.log("🔄 Starting secondary database export...");
      const secondaryExportResult = exportToSecondaryDatabase(allData, summaryStats);
      if (secondaryExportResult.success) {
        Logger.log(`✅ Secondary export completed: ${secondaryExportResult.rowsExported} rows exported to ${secondaryExportResult.targetSpreadsheet}`);
        summaryStats.secondaryExport = secondaryExportResult;
      } else {
        Logger.log(`❌ Secondary export failed: ${secondaryExportResult.message}`);
        summaryStats.secondaryExport = secondaryExportResult;
      }
      
      return;
    }
  } catch (error) {
    Logger.log(`❌ Error in F2DataDumpDirectPrint: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  } finally {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (clearedData && Object.keys(clearedData).length > 0) {
        restoreClearedData(ss, clearedData);
      }
      // Mark the processed email thread as read just before sending the summary email
      if (processedThread) {
        processedThread.markRead();
        Logger.log(`📧 Marked thread as read: ${processedThread.getFirstMessageSubject ? processedThread.getFirstMessageSubject() : ''}`);
      }
      // Send summary email only if we found and processed a matching email
      if (foundMatchingEmail) {
        try {
          const subject = '✅ Barcode Automation Completed Successfully';
          let body = `Barcode automation completed successfully.\n\n` +
            `Dev Mode: ${devMode}\n` +
            `Source sheet rows: ${summaryStats.sourceRows}\n` +
            `Source barcode count: ${summaryStats.sourceBarcodeCount}\n` +
            `Processed rows written: ${summaryStats.totalRows}\n` +
            `Processed barcode count: ${summaryStats.barcodeCount}\n`;
          
          // Add secondary export information if available
          if (summaryStats.secondaryExport) {
            if (summaryStats.secondaryExport.success) {
              body += `\nSecondary Database Export:\n` +
                `✅ Success: ${summaryStats.secondaryExport.rowsExported} rows exported to ${summaryStats.secondaryExport.targetSpreadsheet}\n` +
                `Target Sheet: ${summaryStats.secondaryExport.targetSheet}\n`;
            } else {
              body += `\nSecondary Database Export:\n` +
                `❌ Failed: ${summaryStats.secondaryExport.message}\n`;
            }
          }
          MailApp.sendEmail({
            to: "Share@keslowcamera.com",
            subject,
            body
          });
          Logger.log('✅ Summary email sent.');
        } catch (emailError) {
          Logger.log(`❌ Error sending summary email: ${emailError.toString()}`);
        }
      }
    } catch (restoreError) {
      Logger.log(`❌ Error restoring cleared data: ${restoreError.toString()}`);
    }
  }
} 