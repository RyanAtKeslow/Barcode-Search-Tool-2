/**
 * Update Shipping Database
 * 
 * This script:
 * 1. Checks Gmail for emails with title "Assets Excel Export for Google"
 * 2. Downloads the Excel attachment and converts it to Google Sheets
 * 3. Cross-references data against the "Database" sheet
 * 4. Updates the Database sheet while preserving columns H, I, J data
 */

function updateShippingDatabase() {
  Logger.log("🚀 Starting Update Shipping Database process");
  
  try {
    const maxRetries = 3;
    const retryDelay = 5000;
    
    // Step 1: Search for emails with the specific title (read or unread)
    Logger.log("📧 Searching for emails with title 'Assets Excel Export for Google'");
    const threads = GmailApp.search('subject:"Assets Excel Export for Google"');
    
    if (!threads.length) {
      Logger.log("📭 No matching email found. Exiting quietly.");
      return;
    }
    
    Logger.log(`📧 Found ${threads.length} matching email threads`);
    
         // Step 2: Process email threads and attachments
     Logger.log("📧 Processing email threads...");
     const convertedSheetId = processEmailThreads(threads);
    
    if (!convertedSheetId) {
      Logger.log("❌ Failed to process email attachment");
      return;
    }
    
    Logger.log(`✅ Successfully converted attachment to Google Sheet: ${convertedSheetId}`);
    
    // Step 3: Cross-reference and update Database sheet
    Logger.log("🔄 Starting Database sheet update process");
    updateDatabaseSheet(convertedSheetId);
    
    Logger.log("✅ Update Shipping Database process completed successfully");
    
  } catch (error) {
    Logger.log(`❌ Error in updateShippingDatabase: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Process email threads and attachments following the proven F2 pattern
 * @param {GmailThread[]} threads - Array of Gmail threads
 * @returns {string|null} The ID of the converted Google Sheet
 */
function processEmailThreads(threads) {
  Logger.log("📧 Processing email threads...");
  
  try {
    const startTime = new Date().getTime();
    const MAX_RUNTIME_MS = 4 * 60 * 1000; // 4 minutes
    
    for (const thread of threads) {
      Logger.log("📧 Getting messages from thread...");
      const messages = thread.getMessages();
      Logger.log(`📧 Found ${messages.length} messages in thread`);
      
      if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
        Logger.log('⏳ Approaching timeout, exiting early to avoid session hang.');
        return null;
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
            title: `Converted_${uploadedFile.getName()}_${new Date().getTime()}`,
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
          
          // Clean up the original uploaded file
          DriveApp.getFileById(uploadedFile.getId()).setTrashed(true);
          Logger.log("🗑️ Cleaned up original uploaded file");
          
          return convertedFile.id;
          
        } catch (error) {
          Logger.log(`❌ Error processing file: ${error.toString()}`);
          Logger.log(`Stack trace: ${error.stack}`);
          throw error;
        }
      }
    }
    
    Logger.log("📭 No suitable Excel attachments found in any messages");
    return null;
    
  } catch (error) {
    Logger.log(`❌ Error processing email threads: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return null;
  }
}

/**
 * Update the Database sheet with new data while preserving columns H, I, J
 * @param {string} convertedSheetId - The ID of the converted Google Sheet
 */
function updateDatabaseSheet(convertedSheetId) {
  Logger.log("🔄 Starting Database sheet update process");
  
  try {
    // Open the converted sheet
    Logger.log("🔍 Opening converted sheet...");
    const convertedSheet = SpreadsheetApp.openById(convertedSheetId);
    const convertedData = convertedSheet.getActiveSheet().getDataRange().getValues();
    
    Logger.log(`📊 Converted sheet has ${convertedData.length} rows`);
    
    // Open the Database sheet (assuming it's in the current spreadsheet)
    Logger.log("🔍 Opening Database sheet...");
    const databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
    
    if (!databaseSheet) {
      throw new Error("❌ Database sheet not found in current spreadsheet");
    }
    
    const databaseData = databaseSheet.getDataRange().getValues();
    Logger.log(`📊 Database sheet has ${databaseData.length} rows`);
    
    // Create a map of existing database records by barcode (column C)
    Logger.log("🗺️ Creating database map by barcode...");
    const databaseMap = new Map();
    const headerRow = databaseData[0];
    
    for (let i = 1; i < databaseData.length; i++) {
      const barcode = databaseData[i][2]; // Column C (index 2)
      if (barcode) {
        databaseMap.set(barcode.toString(), {
          rowIndex: i,
          data: databaseData[i]
        });
      }
    }
    
    Logger.log(`📊 Created database map with ${databaseMap.size} existing barcodes`);
    
    // Process converted data to find only new barcodes (skip header row)
    Logger.log("🔄 Processing converted data to identify new barcodes...");
    const newRows = [];
    let existingCount = 0;
    let kcSkippedCount = 0;
    
    for (let i = 1; i < convertedData.length; i++) {
      const convertedRow = convertedData[i];
      const barcode = convertedRow[2]; // Column C
      
      if (!barcode) continue;
      
      const barcodeStr = barcode.toString();
      
      if (databaseMap.has(barcodeStr)) {
        // Barcode already exists - skip it (preserve existing data)
        existingCount++;
      } else {
        // Check if barcode starts with "KC" - if so, skip adding it
        if (barcodeStr.startsWith('KC')) {
          Logger.log(`⚠️ Skipping KC barcode: ${barcodeStr} (KC barcodes not added to database)`);
          kcSkippedCount++;
        } else {
          // New barcode - add as new row with empty H, I, J columns
          const newRow = [...convertedRow.slice(0, 7), '', '', ''];
          newRows.push(newRow);
        }
      }
    }
    
    Logger.log(`📊 Analysis complete:`);
    Logger.log(`  - ${existingCount} barcodes already exist (will be preserved as-is)`);
    Logger.log(`  - ${newRows.length} new barcodes to add`);
    Logger.log(`  - ${kcSkippedCount} KC barcodes skipped (not added to database)`);
    Logger.log(`  - ${databaseMap.size} existing barcodes not found in converted data (will remain in database)`);
    
    
    // Append new rows (much faster than individual updates)
    if (newRows.length > 0) {
      Logger.log("➕ Appending new rows to database...");
      const lastRow = databaseSheet.getLastRow();
      
      // Add all new rows at once
      databaseSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      Logger.log(`✅ Appended ${newRows.length} new rows`);
    }
    
    // Sort the entire sheet by column B (preserving header row)
    Logger.log("🔄 Sorting entire database by column B (preserving headers)...");
    const totalRows = databaseSheet.getLastRow();
    if (totalRows > 1) {
      // Sort from row 2 downward to preserve header row
      const sortRange = databaseSheet.getRange(2, 1, totalRows - 1, databaseSheet.getLastColumn());
      sortRange.sort({column: 2, ascending: true}); // Sort by column B
      Logger.log("✅ Database sorted by column B with headers preserved");
    }
    
    Logger.log("✅ Database sheet update completed successfully");
    
    // Clean up - delete the temporary converted sheet
    Logger.log("🗑️ Cleaning up temporary converted sheet...");
    DriveApp.getFileById(convertedSheetId).setTrashed(true);
    Logger.log("✅ Cleaned up temporary converted sheet");
    
  } catch (error) {
    Logger.log(`❌ Error updating Database sheet: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Manual trigger function for testing
 */
function testUpdateShippingDatabase() {
  Logger.log("🧪 Running test of Update Shipping Database");
  updateShippingDatabase();
} 