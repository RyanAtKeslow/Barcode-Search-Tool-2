function flawlessDataDump() {
  try {
    const writeChunkSize = 10000;
    const maxRetries = 3;
    const retryDelay = 5000;
    const recipient = "Owen@keslowcamera.com, ryan@keslowcamera.com";
    const devMode = true; // Set to true for development mode
    
    Logger.log("🔍 Getting active spreadsheet...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Could not access active spreadsheet");
    }
    const targetSheetName = 'Barcode Dictionary';
    Logger.log("🔍 About to get sheet by name...");
    const sheet = ss.getSheetByName(targetSheetName);
    Logger.log("🔍 getSheetByName finished.");
    if (!sheet) {
      Logger.log("🔍 Sheet not found, about to insert sheet...");
      const newSheet = ss.insertSheet(targetSheetName);
      Logger.log("🔍 insertSheet finished.");
    }
    Logger.log("✅ Target sheet ready.");

    let allData;
    if (devMode) {
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
      // Continue with the rest of the script as normal
      Logger.log(`📊 Source sheet has ${sourceSheet.getLastRow()} rows`);
      Logger.log("📊 Reading source data...");
      const sourceLastRow = sourceSheet.getLastRow();
      const sourceBarcodeCount = sourceSheet.getRange("C2:C" + sourceLastRow).getValues().flat().filter(cell => cell).length;
      Logger.log(`📊 Source sheet statistics:\n- Total rows: ${sourceLastRow}\n- Data rows: ${sourceLastRow - 1}\n- Raw barcode count: ${sourceBarcodeCount}`);
      Logger.log("🔄 Processing and formatting data...");
      const processedData = sortFlawlessDataAutomationMode(allData);
      if (!processedData || processedData.length === 0) {
        Logger.log("⚠️ No data returned from sortFlawlessDataAutomationMode.");
        return;
      }
      Logger.log(`📊 Processed data statistics:\n- Total rows: ${processedData.length}\n- Header row: ${processedData[0].join(', ')}`);
      Logger.log("🔍 Verifying processed data...");
      Logger.log(`Processed data: ${processedData.map(row => row[6]).join(', ')}`);
      Logger.log("📝 Writing processed data to Temp Sheet...");
      const tempSheet = writeToTempSheet(processedData, ss);
      return;
    } else {
      Logger.log("📧 Searching for unread emails...");
      const threads = GmailApp.search('is:unread subject:"Assets Excel Export for Google"');
      if (!threads.length) {
        Logger.log("📭 No matching unread email found. Exiting quietly.");
        return;
      }
      Logger.log(`📧 Found ${threads.length} matching email threads`);
      Logger.log("📡 Barcode automation started.");
      MailApp.sendEmail({
        to: recipient,
        subject: "📡 Barcode Automation Started",
        body: `The barcode automation started running at ${new Date().toLocaleString()}.`
      });
      Logger.log("🚀 Starting flawlessDataDump function...");
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
            break;
          } catch (error) {
            Logger.log(`❌ Error processing file: ${error.toString()}`);
            Logger.log(`Stack trace: ${error.stack}`);
            throw error;
          }
        }
        if (allData) break;
      }
    }

    // Continue with the rest of the script using allData...
    // ... existing code for processing, writing to temp, etc. ...

  } catch (error) {
    Logger.log(`❌ Error in flawlessDataDump: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

function copySheetData(sourceSheet, targetSheet, sourceRange, targetRange) {
  // Get the source range
  const range = sourceSheet.getRange(sourceRange);
  
  // Copy to target range
  range.copyTo(targetSheet.getRange(targetRange), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  // Copy formatting if needed
  range.copyTo(targetSheet.getRange(targetRange), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function getUniqueKey(row) {
  // Category (0), Location (1), Status (2), Equipment Name (3), Owner (4), UUID (5)
  return [row[0], row[1], row[2], row[3], row[4], row[5]].join('||');
}

function logChangesToAnalytics(barcodesToAdd, barcodesToRemove) {
  Logger.log("📊 Logging changes to Analytics sheet...");
  const analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Analytics");
  const timestamp = new Date();
  const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm:ss");
  
  // Set timestamp
  analyticsSheet.getRange("A8").setValue(formattedDate);
  
  // Set barcodes added
  analyticsSheet.getRange("A9").setValue("Barcodes added");
  analyticsSheet.getRange("B9").setValue(barcodesToAdd.length);
  analyticsSheet.getRange("C9").setValue(barcodesToAdd.map(row => row[6]).join(', '));
  
  // Set barcodes removed
  analyticsSheet.getRange("A10").setValue("Barcodes Removed");
  analyticsSheet.getRange("B10").setValue(barcodesToRemove.length);
  analyticsSheet.getRange("C10").setValue(barcodesToRemove.map(row => row[6]).join(', '));
  
  Logger.log("✅ Changes logged to Analytics sheet.");
}

function writeToTempSheet(processedData, ss) {
  Logger.log("📝 Writing processed data to Temp Sheet...");
  let tempSheet = ss.getSheetByName('Temp Sheet');
  if (!tempSheet) {
    tempSheet = ss.insertSheet('Temp Sheet');
  } else {
    tempSheet.clearContents();
    tempSheet.clearFormats();
  }
  // Use the chunk size from the top of the script
  var writeChunkSize = typeof writeChunkSize !== 'undefined' ? writeChunkSize : 10000;
  for (let i = 0; i < processedData.length; i += writeChunkSize) {
    const chunk = processedData.slice(i, i + writeChunkSize);
    tempSheet.getRange(i + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
    Logger.log(`✅ Wrote rows ${i + 1} to ${i + chunk.length} to Temp Sheet.`);
  }
  Logger.log("✅ Processed data written to Temp Sheet.");
  return tempSheet;
}