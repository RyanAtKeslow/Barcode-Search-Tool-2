/**
 * Ship Prep Bay - Prep Bay Shipping and Data Processing Script
 * 
 * This script processes prep bay data for shipping, handling Lost & Found items,
 * updating analytics, and creating comprehensive CSV exports.
 * 
 * Step-by-step process:
 * 1. Identifies the current user using email lookup
 * 2. Searches for username matches in row 2 of the prep bays sheet
 * 3. Handles multiple username matches with user selection dialog
 * 4. Determines column positions (add barcodes, drop barcodes, item names) relative to username
 * 5. Calculates prep bay number from column position
 * 6. Checks for export tags and warns if data hasn't been exported
 * 7. Processes Lost & Found items based on status keywords in item names
 * 8. Updates analytics counters for barcodes and Lost & Found items
 * 9. Filters out empty rows and export tag rows
 * 10. Creates CSV content with headers and all data
 * 11. Generates timestamped filename with prep bay number
 * 12. Saves CSV file to "Prep Bay Scans" folder in Google Drive
 * 13. Displays download link with data clearing functionality
 * 14. Sends status update to database with "Shipped" status
 * 15. Clears bay data and resets background colors
 * 
 * Lost & Found Processing:
 * - Identifies items with status keywords (Disposed, Repair, Lost, Inactive, Sale, Pending QC)
 * - Extracts consigner information from item names
 * - Updates quantities for existing items or adds new ones
 * - Handles duplicate barcodes within the same batch
 * 
 * Analytics Updates:
 * - Counts unique barcodes from both add and drop columns
 * - Counts Lost & Found items based on keyword detection
 * - Updates total counters in Analytics sheet
 * 
 * Features:
 * - Comprehensive Lost & Found processing
 * - Analytics tracking and updates
 * - Export tag validation and warnings
 * - Data clearing after successful export
 * - Database status updates
 * - Google Drive integration
 */
// Helper function to normalize barcodes by removing pipes and trimming
function normalizeBarcode(barcode) {
  if (!barcode) return '';
  return barcode.toString().trim();
}

// Helper function to split pipe-delimited barcodes and normalize each one
function splitAndNormalizeBarcodes(barcodeString) {
  if (!barcodeString) return [];
  return barcodeString.toString()
    .split('|')
    .map(b => normalizeBarcode(b))
    .filter(b => b.length > 0);  // Remove empty strings
}

// Helper function to extract job name from username cell
function extractJobName(usernameCellValue, username) {
  if (!usernameCellValue) return '';
  // Remove the username from the cell value and trim
  return usernameCellValue.toString()
    .replace(username, '')
    .trim();
}

/**
 * Processes items for Lost & Found sheet based on keywords in item names
 * @param {Array} addItemNames - Array of add item names
 * @param {Array} dropItemNames - Array of drop item names
 * @param {Array} addBarcodes - Array of add barcodes
 * @param {Array} dropBarcodes - Array of drop barcodes
 * @param {string} jobInfo - Job info from username cell
 * @returns {Array} Array of Lost & Found items to append
 */
function processLostAndFoundItems(addItemNames, dropItemNames, addBarcodes, dropBarcodes, jobInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lostAndFoundSheet = ss.getSheetByName("Lost & Found");
  
  // Get existing data from Lost & Found to check for duplicates and update quantities
  const existingData = lostAndFoundSheet.getRange("A:D").getValues();
  const existingBarcodeMap = new Map(); // Map barcode to row index
  
  // Build map of existing barcodes and their row positions
  for (let i = 0; i < existingData.length; i++) {
    const barcode = existingData[i][0];
    if (barcode && barcode.toString().trim() !== "") {
      existingBarcodeMap.set(barcode.toString(), i + 1); // +1 for 1-based row indexing
    }
  }
  
  const lostAndFoundItems = [];
  const quantityUpdates = []; // Track quantity updates for existing items
  const keywordsToRemove = ["Disposed", "Repair", "Lost", "Inactive", "Sale", "Pending QC"];
  
  // Process both add and drop items
  const processItems = (itemNames, barcodes) => {
    for (let i = 0; i < itemNames.length; i++) {
      const itemName = itemNames[i][0];
      const barcode = barcodes[i][0];
      
      if (!itemName || !barcode) continue;
      
      // Check for keywords in item name
      let statusText = "";
      let modifiedItemName = itemName;
      
      for (const keyword of keywordsToRemove) {
        const regex = new RegExp(`\\b${keyword}\\b(?=\\s*\\||$)`, "i");
        if (regex.test(itemName)) {
          statusText = keyword.charAt(0) + keyword.slice(1).toLowerCase();
          modifiedItemName = itemName.replace(regex, "").replace(/\s+\|/, "|").trim();
          break;
        }
      }
      
      if (statusText) {
        // Extract consigner from item name if present
        let consigner = "";
        if (modifiedItemName.includes("|")) {
          const parts = modifiedItemName.split("|");
          modifiedItemName = parts[0].trim();
          consigner = parts[1].trim();
        } else {
          modifiedItemName = modifiedItemName.trim();
          consigner = "Keslow"; // Default consigner
        }
        
        // Clean up dangling pipes or extra spaces
        modifiedItemName = modifiedItemName.replace(/\|\s*$/, "").trim();
        
        // Check if barcode already exists in Lost & Found
        if (existingBarcodeMap.has(barcode.toString())) {
          const existingValue = existingBarcodeMap.get(barcode.toString());
          
          // Only process if it's a valid row number (not "pending")
          if (typeof existingValue === 'number') {
            // Increment quantity for existing item
            const currentQuantity = lostAndFoundSheet.getRange(existingValue, 4).getValue() || 0;
            quantityUpdates.push({
              row: existingValue,
              newQuantity: currentQuantity + 1,
              jobInfo: jobInfo
            });
          } else {
            // This is a "pending" item from this batch, increment its quantity in the pending items
            const pendingIndex = lostAndFoundItems.findIndex(item => item[0] === barcode.toString());
            if (pendingIndex !== -1) {
              lostAndFoundItems[pendingIndex][3] += 1; // Increment quantity (column D)
            }
          }
        } else {
          // Add new item to Lost & Found
          lostAndFoundItems.push([barcode, modifiedItemName, statusText, 1, jobInfo, "", consigner]);
          // Add to map to track for subsequent duplicates in this batch
          existingBarcodeMap.set(barcode.toString(), "pending");
        }
      }
    }
  };
  
  // Process both add and drop items
  processItems(addItemNames, addBarcodes);
  processItems(dropItemNames, dropBarcodes);
  
  // Update quantities for existing items
  quantityUpdates.forEach(update => {
    lostAndFoundSheet.getRange(update.row, 4).setValue(update.newQuantity);
    lostAndFoundSheet.getRange(update.row, 5).setValue(update.jobInfo); // Overwrite column E
  });
  
  // Append new lost and found items if any found
  if (lostAndFoundItems.length > 0) {
    const lastDataRowLF = lostAndFoundSheet.getRange("B:B").getValues().filter(row => row[0].toString().trim() !== "").length;
    const targetRangeLF = lostAndFoundSheet.getRange(lastDataRowLF + 1, 1, lostAndFoundItems.length, 7);
    targetRangeLF.setValues(lostAndFoundItems);
    
    const checkboxRange = lostAndFoundSheet.getRange(lastDataRowLF + 1, 6, lostAndFoundItems.length, 1);
    checkboxRange.insertCheckboxes();
  }
  
  return lostAndFoundItems;
}

/**
 * Updates the analytics sheet with the total barcode count
 * @param {number} barcodeCount - Number of barcodes to add to the total
 * @returns {number} The new total after adding the count
 */
function updateAnalyticsBarcodeCount(barcodeCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analyticsSheet = ss.getSheetByName("Analytics");
    
    // Get current value from Z2
    const currentTotal = analyticsSheet.getRange("Z2").getValue() || 0;
    
    // Calculate new total
    const newTotal = currentTotal + barcodeCount;
    
    // Update Z2 with new total
    analyticsSheet.getRange("Z2").setValue(newTotal);
    return newTotal;
  } catch (error) {
    Logger.log(`❌ Error in updateAnalyticsBarcodeCount: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Updates the analytics sheet with the total Lost & Found count
 * @param {number} lostAndFoundCount - Number of Lost & Found items to add to the total
 * @returns {number} The new total after adding the count
 */
function updateAnalyticsLostAndFoundCount(lostAndFoundCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analyticsSheet = ss.getSheetByName("Analytics");
    
    // Get current value from AA2
    const currentTotal = analyticsSheet.getRange("AA2").getValue() || 0;
    
    // Calculate new total
    const newTotal = currentTotal + lostAndFoundCount;
    
    // Update AA2 with new total
    analyticsSheet.getRange("AA2").setValue(newTotal);
    return newTotal;
  } catch (error) {
    Logger.log(`❌ Error in updateAnalyticsLostAndFoundCount: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

function ShipPrepBay() {
  try {
    // Get user's email and extract username
    const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prepBaysSheet = ss.getSheetByName("prep bays");
    const ui = SpreadsheetApp.getUi();
    
    // Search for username in prep bays sheet using two-step process
    const usernameMatches = findUsernameInRow2(prepBaysSheet);
    
    let usernameCell;
    
    // Handle multiple matches
    if (usernameMatches.length > 1) {
      const selectedMatch = setSelectedMatch(usernameMatches);
      if (!selectedMatch) {
        throw new Error('No match was selected');
      }
      usernameCell = prepBaysSheet.getRange(selectedMatch.cellA1);
    } else if (usernameMatches.length === 1) {
      usernameCell = prepBaysSheet.getRange(usernameMatches[0].cellA1);
    } else {
      ui.alert(
        'Name Not Found',
        'Please input your name as it\'s shown in your keslow email into the name fields, followed by your Job name and Contract #',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Define column locations based on usernameCell
    const usernameCol = usernameCell.getColumn();
    const addBarcodesCol = usernameCol - 1;
    const addItemNamesCol = usernameCol;
    const dividerCol = usernameCol + 1;
    const dropBarcodesCol = usernameCol + 2;
    const dropItemNamesCol = usernameCol + 3;
    
    // Calculate prep bay number from column
    // B2 = 1, H2 = 2, N2 = 3, etc.
    // Each prep bay is 6 columns apart (B->H->N)
    const prepBayNumber = Math.floor((usernameCol - 2) / 6) + 1;
    
    // Check for export tags by scanning up from the bottom
    const lastRow = prepBaysSheet.getLastRow();
    const addBarcodesData = prepBaysSheet.getRange(4, addBarcodesCol, lastRow - 3, 1).getValues();
    const dropBarcodesData = prepBaysSheet.getRange(4, dropBarcodesCol, lastRow - 3, 1).getValues();
    
    let hasAddExportTag = false;
    let hasDropExportTag = false;
    let lastAddBarcode = -1;
    let lastDropBarcode = -1;
    let lastAddExportTag = -1;
    let lastDropExportTag = -1;
    
    // Scan add barcodes column
    for (let i = addBarcodesData.length - 1; i >= 0; i--) {
      const value = addBarcodesData[i][0];
      if (value) {
        if (value.toString().includes("Above was exported @")) {
          lastAddExportTag = i;
          hasAddExportTag = true;
          break;
        } else {
          lastAddBarcode = i;
        }
      }
    }
    
    // Scan drop barcodes column
    for (let i = dropBarcodesData.length - 1; i >= 0; i--) {
      const value = dropBarcodesData[i][0];
      if (value) {
        if (value.toString().includes("Above was exported @")) {
          lastDropExportTag = i;
          hasDropExportTag = true;
          break;
        } else {
          lastDropBarcode = i;
        }
      }
    }
    
    // Only show warning if there are barcodes after the last export tag in either column
    const needsWarning = (lastAddBarcode > lastAddExportTag && lastAddBarcode !== -1) || 
                        (lastDropBarcode > lastDropExportTag && lastDropBarcode !== -1);
    
    if (needsWarning) {
      const response = ui.alert(
        'Export Tags Missing',
        `You may have forgotten to export your barcodes in Prep Bay ${prepBayNumber}! Continue anyway?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.NO) {
        return;
      }
    }
    
    // Get data from row 4 onwards
    const addBarcodesDataRange = prepBaysSheet.getRange(4, addBarcodesCol, lastRow - 3, 1);
    const addItemNamesDataRange = prepBaysSheet.getRange(4, addItemNamesCol, lastRow - 3, 1);
    const dropBarcodesDataRange = prepBaysSheet.getRange(4, dropBarcodesCol, lastRow - 3, 1);
    const dropItemNamesDataRange = prepBaysSheet.getRange(4, dropItemNamesCol, lastRow - 3, 1);
    
    const addBarcodes = addBarcodesDataRange.getValues();
    const addItemNames = addItemNamesDataRange.getValues();
    const dropBarcodes = dropBarcodesDataRange.getValues();
    const dropItemNames = dropItemNamesDataRange.getValues();
    
    // Process Lost & Found items
    const jobInfo = usernameCell.getValue().toString().trim();
    const lostAndFoundItems = processLostAndFoundItems(addItemNames, dropItemNames, addBarcodes, dropBarcodes, jobInfo);
    
    // Filter out export tags and empty cells, combine the data
    const filteredData = [];
    for (let i = 0; i < addBarcodes.length; i++) {
      const addBarcode = addBarcodes[i][0];
      const addItemName = addItemNames[i][0];
      const dropBarcode = dropBarcodes[i][0];
      const dropItemName = dropItemNames[i][0];
      
      // Skip if all cells in this row are empty
      if (!addBarcode && !dropBarcode) continue;
      
      // Skip export tag rows
      if ((addBarcode && addBarcode.toString().includes("Above was exported @")) || 
          (dropBarcode && dropBarcode.toString().includes("Above was exported @"))) {
        continue;
      }
      
      filteredData.push([
        addBarcode || '',      // Add barcode
        addItemName || '',     // Add item name
        dropBarcode || '',     // Drop barcode
        dropItemName || ''     // Drop item name
      ]);
    }
    
    // Count unique barcodes from both add and drop columns
    const uniqueBarcodes = new Set();
    const lostAndFoundBarcodes = new Set();
    const keywordsToCheck = ["Disposed", "Repair", "Lost", "Inactive", "Sale", "Pending QC"];
    
    filteredData.forEach(row => {
      const addBarcode = row[0];
      const addItemName = row[1];
      const dropBarcode = row[2];
      const dropItemName = row[3];
      
      // Add to unique barcodes set
      if (addBarcode) uniqueBarcodes.add(addBarcode);
      if (dropBarcode) uniqueBarcodes.add(dropBarcode);
      
      // Check for Lost & Found keywords in item names
      const hasKeyword = (itemName) => {
        return keywordsToCheck.some(keyword => 
          itemName && itemName.toString().toLowerCase().includes(keyword.toLowerCase())
        );
      };
      
      if (addBarcode && hasKeyword(addItemName)) {
        lostAndFoundBarcodes.add(addBarcode);
      }
      if (dropBarcode && hasKeyword(dropItemName)) {
        lostAndFoundBarcodes.add(dropBarcode);
      }
    });
    
    // Update analytics with barcode counts
    const barcodeCount = uniqueBarcodes.size;
    const lostAndFoundCount = lostAndFoundBarcodes.size;
    
    const newBarcodeTotal = updateAnalyticsBarcodeCount(barcodeCount);
    const newLostAndFoundTotal = updateAnalyticsLostAndFoundCount(lostAndFoundCount);
    
    Logger.log(`updateAnalyticsBarcodeCount: +${barcodeCount}, new total ${newBarcodeTotal}`);
    Logger.log(`updateAnalyticsLostAndFoundCount: +${lostAndFoundCount}, new total ${newLostAndFoundTotal}`);
    
    // Create CSV content with headers
    const headers = ['Add Barcode', 'Add Item Name', 'Drop Barcode', 'Drop Item Name'];
    const csvContent = [
      headers.join(','),
      ...filteredData.map(row => row.join(','))
    ].join('\n');
    
    // Create filename with timestamp using the username cell value
    const timestamp = getConsistentTimestamp();
    const cellValue = usernameCell.getValue();
    const filename = `${cellValue} - Shipped (${timestamp})_Bay${prepBayNumber}.csv`;
    
    // Save to Shared Drive
    saveToSharedDrive('Prep Bay Scans', filename, csvContent);
    
    // Also save to user's personal Drive as backup
    const file = DriveApp.createFile(filename, csvContent, MimeType.CSV);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Create or get the 'Prep Bay Scans' folder in the user's Drive
    let folders = DriveApp.getFoldersByName('Prep Bay Scans');
    let prepBayFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Prep Bay Scans');
    prepBayFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Remove from root

    // Show success message with download link
    const html = HtmlService.createHtmlOutput(`
      <p>✅ Successfully exported ${filteredData.length} items!</p>
      <p>The file has been saved to your Google Drive and is ready to download.</p>
      <a href="${file.getUrl()}" target="_blank" onclick="google.script.run.clearDataAfterDownload('${usernameCell.getA1Notation()}', ${addBarcodesCol}, ${dropBarcodesCol})" style="
        display: inline-block;
        padding: 10px 20px;
        background-color: #4285f4;
        color: white;
        text-decoration: none;
        border-radius: 4px;
        margin-top: 10px;
      ">Download CSV File</a>
    `)
    .setWidth(400)
    .setHeight(200);
    
    ui.showModalDialog(html, 'Export Complete');
    
    // Get job name from username cell value
    const jobName = cellValue.toString().trim();
    
    // Send digital cameras to database with Shipped status
    // Pass the filtered barcodes array, username, and job name
    const barcodesArray = Array.from(new Set(filteredData.map(row => row[0]))); // Get unique Add barcodes
    SendStatus("Shipped", barcodesArray, username, jobName, userEmail);
    
    // Clear the data from add barcodes, drop barcodes, and username cell
    addBarcodesDataRange.clearContent();
    dropBarcodesDataRange.clearContent();
    usernameCell.clearContent();
    
    // Reset background colors to white
    addBarcodesDataRange.setBackground("#ffffff");
    dropBarcodesDataRange.setBackground("#ffffff");
    usernameCell.setBackground("#ffffff");
    
  } catch (error) {
    Logger.log(`❌ Error in ShipPrepBay: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

function validatePasswordAndSetLocation(password, prepBay) {
  if (password !== "password") {
    throw new Error("Invalid password");
  }
  
  // Calculate cell location based on prep bay number
  const col = prepBay * 5 + 3;
  const row = 2;
  
  // Store the location for use in the main function
  PropertiesService.getScriptProperties().setProperty('prepBayLocation', `${col},${row}`);
  
  // Run the main export function again
  ShipPrepBay();
}

function showPasswordDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <form id="passwordForm">
      <div style="margin-bottom: 10px;">
        <label for="password">Password:</label>
        <input type="password" id="password" name="password" style="margin-left: 10px;">
      </div>
      <div style="margin-bottom: 10px;">
        <label for="prepBay">Prep Bay:</label>
        <select id="prepBay" name="prepBay" style="margin-left: 10px;">
          ${Array.from({length: 22}, (_, i) => `<option value="${i + 1}">Prep Bay ${i + 1}</option>`).join('')}
        </select>
      </div>
      <input type="button" value="Submit" onclick="submitForm()" style="
        padding: 5px 15px;
        background-color: #4285f4;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
      ">
    </form>
    <script>
      function submitForm() {
        const password = document.getElementById('password').value;
        const prepBay = document.getElementById('prepBay').value;
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler((error) => alert(error))
          .validatePasswordAndSetLocation(password, prepBay);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(200);
  
  ui.showModalDialog(html, 'Enter Password');
}

// saveToSharedDrive function is now centralized in Data Model.js

function clearDataAfterDownload(usernameCellA1, addBarcodesCol, dropBarcodesCol) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("prep bays");
    
    // Get the ranges to clear
    const usernameCell = sheet.getRange(usernameCellA1);
    const lastRow = sheet.getLastRow();
    const addBarcodesRange = sheet.getRange(4, addBarcodesCol, lastRow - 3, 1);
    const dropBarcodesRange = sheet.getRange(4, dropBarcodesCol, lastRow - 3, 1);
    
    // Clear the content
    addBarcodesRange.clearContent();
    dropBarcodesRange.clearContent();
    usernameCell.clearContent();
    
    // Reset background colors to white
    addBarcodesRange.setBackground("#ffffff");
    dropBarcodesRange.setBackground("#ffffff");
    usernameCell.setBackground("#ffffff");
    
  } catch (error) {
    Logger.log(`❌ Error in clearDataAfterDownload: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
} 