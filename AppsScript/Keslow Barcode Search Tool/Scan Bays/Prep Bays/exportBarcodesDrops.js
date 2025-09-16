/**
 * Export Barcodes Drops - Prep Bay Drop Export Script
 * 
 * This script exports drop barcode data from prep bays to CSV files
 * with duplicate detection and status tracking for returned items.
 * 
 * Step-by-step process:
 * 1. Identifies the current user using email lookup
 * 2. Searches for username matches in row 2 of the prep bays sheet
 * 3. Handles multiple username matches with user selection dialog
 * 4. Determines column positions (add barcodes, drop barcodes) relative to username
 * 5. Checks for duplicate barcodes between add and drop columns
 * 6. Highlights duplicate cells in yellow if found and stops execution
 * 7. Scans for export tags to identify new data since last export
 * 8. Extracts barcodes from the drop column (right of username)
 * 9. Filters out empty cells and export tag rows
 * 10. Creates CSV content with barcode data
 * 11. Generates timestamped filename with prep bay number
 * 12. Saves CSV file to "Prep Bay Scans" folder in Google Drive
 * 13. Displays download link and success message
 * 14. Highlights exported barcodes in light green
 * 15. Adds export tag with timestamp
 * 16. Sends status update to database with "Returned" status
 * 
 * Duplicate Detection:
 * - Compares barcodes between add and drop columns
 * - Highlights conflicting cells in yellow
 * - Stops execution until duplicates are resolved
 * - Provides detailed information about conflicts
 * 
 * Export Tracking:
 * - Uses export tags to track previously exported data
 * - Only exports new data since last export
 * - Prevents duplicate exports of the same data
 * 
 * Features:
 * - Comprehensive duplicate detection and handling
 * - Export tracking with timestamp tags
 * - Visual feedback with cell highlighting
 * - Database status updates for returned items
 * - Google Drive integration
 */
function prepBayDropExport() {
  try {
    Logger.log('prepBayDropExport start');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    
    // Get user's email and extract username
    const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
    Logger.log(`User email: ${userEmail}, Username: ${username}`);
    
    // Search for username in row 2 using two-step process
    const usernameMatches = findUsernameInRow2(sheet);
    Logger.log(`Found ${usernameMatches.length} username matches`);
    
    let usernameCell;
    
    // Handle multiple matches
    if (usernameMatches.length > 1) {
      Logger.log('Multiple matches found, showing selection dialog');
      const selectedMatch = setSelectedMatch(usernameMatches);
      if (!selectedMatch) {
        throw new Error('No match was selected');
      }
      usernameCell = sheet.getRange(selectedMatch.cellA1);
    } else if (usernameMatches.length === 1) {
      Logger.log(`Single match found: ${usernameMatches[0].value}`);
      usernameCell = sheet.getRange(usernameMatches[0].cellA1);
    } else {
      Logger.log('No matches found, showing alert...');
      ui.alert(
        'Name Not Found',
        'Please input your name as it\'s shown in your keslow email into the name fields, followed by your Job name and Contract #',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Get the drop barcodes column (two columns right of username)
    const dropBarcodesCol = usernameCell.getColumn() + 2; // two cols right
    Logger.log(`Drop barcodes column: ${dropBarcodesCol}`);
    
    // Get the add barcodes column (three columns left of drop column)
    const addBarcodesCol = dropBarcodesCol - 3;
    Logger.log(`Add barcodes column: ${addBarcodesCol}`);
    
    // Check for duplicates between add and drop columns
    const { addBarcodes, dropBarcodes, duplicates } = checkForDuplicates(sheet, addBarcodesCol, dropBarcodesCol);
    
    // If duplicates found, stop the script with UI alert and highlight cells
    if (duplicates.length > 0) {
      Logger.log(`Found ${duplicates.length} duplicates, stopping script...`);
      
      // Highlight duplicate cells in yellow
      duplicates.forEach(duplicate => {
        // Highlight add column barcode cell
        sheet.getRange(duplicate.addRow, addBarcodesCol).setBackground("#fff2cc");
        // Highlight drop column barcode cell
        sheet.getRange(duplicate.dropRow, dropBarcodesCol).setBackground("#fff2cc");
      });
      
      let message = `Duplicate barcodes found, are these adds or drops?\n\n`;
      message += `Duplicate cells have been highlighted in YELLOW for your review.\n\n`;
      duplicates.forEach((duplicate, index) => {
        message += `${index + 1}. Barcode: ${duplicate.barcode}\n`;
        message += `   Add Column: (Row ${duplicate.addRow}) ${duplicate.addItemName || 'N/A'} \n`;
        message += `   Drop Column: (Row ${duplicate.dropRow}) ${duplicate.dropItemName || 'N/A'} \n\n`;
      });
      
      ui.alert('Duplicate Barcodes Found', message, ui.ButtonSet.OK);
      return;
    }
    
    // Get all values in both columns starting from row 4
    const lastRow = sheet.getLastRow();
    const columnD = sheet.getRange(4, dropBarcodesCol, lastRow - 3, 1).getValues();
    
    let exportTagRow = -1;
    let lastDataRow = -1;
    for (let i = columnD.length - 1; i >= 0; i--) {
      const cellVal = columnD[i][0];
      if (!cellVal) continue;
      const str = cellVal.toString();
      if (lastDataRow === -1 && !str.includes("Above was exported @")) {
        lastDataRow = i + 4;
      }
      if (exportTagRow === -1 && str.includes("Above was exported @")) {
        exportTagRow = i + 4;
      }
      if (lastDataRow !== -1 && exportTagRow !== -1) break;
    }
    
    // If we couldn't find any barcodes, show message and return
    if (lastDataRow === -1) {
      ui.alert('No Data', 'No barcodes found to export.', ui.ButtonSet.OK);
      return;
    }
    
    // Start collecting data from after the last export tag
    const startRow = exportTagRow > -1 ? exportTagRow + 1 : 4;
    
    // If there's no new data after the last export tag, show a message and return
    if (startRow > lastDataRow) {
      ui.alert('No New Data', 'There are no new barcodes to export since the last export.', ui.ButtonSet.OK);
      return;
    }
    
    // Get the range of barcodes to export (only the new ones after the last export tag)
    const barcodeRange = sheet.getRange(startRow, dropBarcodesCol, lastDataRow - startRow + 1, 1);
    const barcodes = barcodeRange.getValues();
    
    // Filter out any empty cells and normalize barcodes
    const filteredData = [];
    for (let i = 0; i < barcodes.length; i++) {
      const barcode = barcodes[i][0];
      if (barcode && barcode.toString().trim() && !barcode.toString().includes("Above was exported @")) {
        // Only include non-empty, non-export-tag cells with trimmed values
        filteredData.push([barcode.toString().trim()]);
      }
    }
    
    // If no data to export after filtering, show a message and return
    if (filteredData.length === 0) {
      ui.alert('No Data', 'No barcodes found to export.', ui.ButtonSet.OK);
      return;
    }
    
    // Create CSV content with just the barcodes
    const csvContent = filteredData.map(row => row.join(',')).join('\n');
    
    // Create filename with timestamp
    const timestamp = getConsistentTimestamp();
    const cellValue = usernameCell.getValue();
    // Derive prep bay number from usernameCell's column
    const prepBayNumberForExport = Math.floor((usernameCell.getColumn() - 3) / 5) + 1;
    Logger.log(`Prep Bay Number: ${prepBayNumberForExport}`);

    // Create filename with timestamp and prep bay number
    const filename = `${cellValue}_${timestamp}_drops_Bay${prepBayNumberForExport}.csv`;
    
    // Save to Shared Drive
    saveToSharedDrive('Prep Bay Scans', filename, csvContent);
    
    // Also save to user's personal Drive as backup
    const file = DriveApp.createFile(filename, csvContent, MimeType.CSV);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    let folders = DriveApp.getFoldersByName('Prep Bay Scans');
    let prepBayFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Prep Bay Scans');
    prepBayFolder.createFile(file);

    // Create a file URL (standard Google Drive link)
    const fileUrl = file.getUrl();
    
    // Show success message with file link
    const html = HtmlService.createHtmlOutput(`
      <p>✅ Successfully exported ${filteredData.length} items!</p>
      <p>The file has been saved to your Google Drive.</p>
      <a href="${fileUrl}" target="_blank" onclick="google.script.run.afterExportComplete()" style="
        display: inline-block;
        padding: 10px 20px;
        background-color: #4285f4;
        color: white;
        text-decoration: none;
        border-radius: 4px;
        margin-top: 10px;
      ">Open File in Google Drive</a>
      <p style="font-size: 12px; color: #666; margin-top: 10px;">Once opened, you can download it using the download button in Google Drive.</p>
    `)
    .setWidth(450)
    .setHeight(220);
    
    ui.showModalDialog(html, 'Export Complete');
    
    // Set background color of only the barcode cells to light green
    barcodeRange.setBackground("#b7e1cd");
    
    // Add export tag in the next row after the last barcode
    const exportTag = `Above was exported @ ${new Date().toLocaleString()}`;
    sheet.getRange(lastDataRow + 1, dropBarcodesCol).setValue(exportTag);
    
    Logger.log(`✅ Drop export complete | items: ${filteredData.length}`);
    
    // Get job name from username cell value
    const jobName = cellValue.toString().trim();
    
    // Send digital cameras to database with Returned status
    // Pass the filtered barcodes array, username, and job name
    const barcodesArray = Array.from(new Set(filteredData.map(row => row[0])));
    const uniqueBarcodes = Array.from(new Set(barcodesArray));
    SendStatus("Returned", uniqueBarcodes, username, jobName, userEmail);
    
  } catch (error) {
    Logger.log(`❌ Error in prepBayDropExport: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

function showPasswordDialog() {
  Logger.log('=== showPasswordDialog START ===');
  Logger.log('Creating password dialog UI...');
  
  const ui = SpreadsheetApp.getUi();
  Logger.log('UI object created');
  
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
        Logger.log('Form submitted with prep bay: ' + prepBay);
        google.script.run
          .withSuccessHandler(() => {
            Logger.log('Password validation successful');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            Logger.log('Password validation failed: ' + error);
            alert(error);
          })
          .validatePasswordAndSetLocation(password, prepBay);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(200);
  
  Logger.log('HTML dialog created, showing dialog...');
  ui.showModalDialog(html, 'Enter Password');
  Logger.log('=== showPasswordDialog END ===');
}

function validatePasswordAndSetLocation(password, prepBay) {
  Logger.log('=== validatePasswordAndSetLocation START ===');
  Logger.log(`Received password and prep bay: ${prepBay}`);
  
  if (password !== "password") {
    Logger.log('❌ Invalid password provided');
    throw new Error("Invalid password");
  }
  Logger.log('✅ Password validated successfully');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  Logger.log('Got active spreadsheet and sheet');
  
  // Calculate cell location based on prep bay number
  // Prep Bay 1 = B2 (col 2), Prep Bay 2 = H2 (col 8), etc.
  const col = (prepBay - 1) * 6 + 2;  // This gives us: 1->2, 2->8, 3->14, etc.
  const row = 2;
  Logger.log(`Calculated column for prep bay ${prepBay}: ${col} (should be ${prepBay === 1 ? 'B' : prepBay === 2 ? 'H' : 'other'})`);
  
  // Get the username cell
  const usernameCell = sheet.getRange(row, col);
  Logger.log(`Username cell location: ${usernameCell.getA1Notation()}`);
  
  // Get user's email and extract username
  const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
  Logger.log(`User email: ${userEmail}, Username: ${username}`);
  
  // Get existing value and append username if needed
  const existingValue = usernameCell.getValue();
  Logger.log(`Existing value in cell: "${existingValue}"`);
  
  if (existingValue) {
    const newValue = `${existingValue}, ${username}`;
    Logger.log(`Appending username to existing value. New value will be: "${newValue}"`);
    usernameCell.setValue(newValue);
    Logger.log('✅ Username appended successfully');
  } else {
    Logger.log(`Setting new username value: "${username}"`);
    usernameCell.setValue(username);
    Logger.log('✅ New username set successfully');
  }
  
  // Wait a moment for the cell value to be set
  Utilities.sleep(1000);
  
  Logger.log('Running main export function again...');
  prepBayDropExport();
  Logger.log('=== validatePasswordAndSetLocation END ===');
}

function afterExportComplete() {
  try {
    // Get user's email and extract username
    const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
    
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Search for username in row 2
    let usernameMatches = [];
    const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (let j = 0; j < row2.length; j++) {
      if (row2[j] && row2[j].toString().toLowerCase().includes(username.toLowerCase())) {
        usernameMatches.push({
          cell: sheet.getRange(2, j + 1),
          value: row2[j]
        });
      }
    }
    
    let usernameCell;
    
    // Handle multiple matches
    if (usernameMatches.length > 1) {
      Logger.log(`Multiple matches found: ${usernameMatches.length}`);
      const html = HtmlService.createHtmlOutput(`
        <form id="matchForm">
          <div style="margin-bottom: 10px;">
            <label>Oops! It looks like you have your name on multiple prep bays. Please select the correct entry:</label>
            <select id="matchSelect" style="margin-top: 10px; width: 100%;">
              ${usernameMatches.map((match, index) => 
                `<option value="${index}">${match.value}</option>`
              ).join('')}
            </select>
          </div>
          <input type="button" value="Select" onclick="submitSelection()" style="
            padding: 5px 15px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          ">
        </form>
        <script>
          function submitSelection() {
            const index = document.getElementById('matchSelect').value;
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler((error) => alert(error))
              .setSelectedMatch(index);
          }
        </script>
      `)
      .setWidth(400)
      .setHeight(200);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Select Correct Entry');
      return;
    } else if (usernameMatches.length === 1) {
      usernameCell = usernameMatches[0].cell;
    } else {
      Logger.log('❌ No username match found in prep bays sheet');
      return;
    }
    
    const jobName = usernameCell.getValue().toString().trim();
    
    // Get barcodes from the sheet
    const barcodes = sheet.getRange(4, 1, sheet.getLastRow() - 3, 1)
      .getValues()
      .map(row => row[0])
      .filter(barcode => barcode && !barcode.toString().includes("Above was exported @"));
    
    // Send digital cameras to database with Returned status
    const uniqueBarcodes = Array.from(new Set(barcodes));
    SendStatus("Returned", uniqueBarcodes, username, jobName, userEmail);
    
  } catch (error) {
    Logger.log(`❌ Error in afterExportComplete: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
} 

// saveToSharedDrive function is now centralized in Data Model.js

function checkForDuplicates(sheet, addColumn, dropColumn) {
  try {
    Logger.log(`Checking for duplicates between columns ${addColumn} and ${dropColumn}`);
    
    // Get all values from both columns starting from row 4
    const lastRow = sheet.getLastRow();
    const addColumnData = sheet.getRange(4, addColumn, lastRow - 3, 2).getValues(); // barcode + item name
    const dropColumnData = sheet.getRange(4, dropColumn, lastRow - 3, 2).getValues(); // barcode + item name
    
    // Find the most recent export tag row for each column
    let addExportTagRow = -1;
    let dropExportTagRow = -1;
    
    // Find add column export tag (scan from bottom)
    for (let i = addColumnData.length - 1; i >= 0; i--) {
      const cellVal = addColumnData[i][0];
      if (cellVal && cellVal.toString().includes("Above was exported @")) {
        addExportTagRow = i + 4;
        break;
      }
    }
    
    // Find drop column export tag (scan from bottom)
    for (let i = dropColumnData.length - 1; i >= 0; i--) {
      const cellVal = dropColumnData[i][0];
      if (cellVal && cellVal.toString().includes("Above was exported @")) {
        dropExportTagRow = i + 4;
        break;
      }
    }
    
    Logger.log(`Add column export tag row: ${addExportTagRow}, Drop column export tag row: ${dropExportTagRow}`);
    
    // Find duplicate barcodes between the two columns (only below export tags)
    const addBarcodes = new Map();
    const dropBarcodes = new Map();
    const duplicates = [];
    
    // Process add column data (only below export tag)
    const addStartRow = addExportTagRow > -1 ? addExportTagRow + 1 : 4;
    for (let i = addStartRow - 4; i < addColumnData.length; i++) {
      const barcode = addColumnData[i][0];
      const itemName = addColumnData[i][1];
      // Ensure barcode is not empty, whitespace-only, or an export tag
      if (barcode && barcode.toString().trim() && !barcode.toString().includes("Above was exported @")) {
        const normalizedBarcode = barcode.toString().trim();
        addBarcodes.set(normalizedBarcode, { row: i + 4, itemName: itemName });
      }
    }
    
    // Process drop column data and find duplicates (only below export tag)
    const dropStartRow = dropExportTagRow > -1 ? dropExportTagRow + 1 : 4;
    for (let i = dropStartRow - 4; i < dropColumnData.length; i++) {
      const barcode = dropColumnData[i][0];
      const itemName = dropColumnData[i][1];
      // Ensure barcode is not empty, whitespace-only, or an export tag
      if (barcode && barcode.toString().trim() && !barcode.toString().includes("Above was exported @")) {
        const normalizedBarcode = barcode.toString().trim();
        dropBarcodes.set(normalizedBarcode, { row: i + 4, itemName: itemName });
        
        // Check if this barcode exists in add column (below export tag)
        if (addBarcodes.has(normalizedBarcode)) {
          duplicates.push({
            barcode: normalizedBarcode,
            addRow: addBarcodes.get(normalizedBarcode).row,
            addItemName: addBarcodes.get(normalizedBarcode).itemName,
            dropRow: i + 4,
            dropItemName: itemName
          });
        }
      }
    }
    
    // Convert maps to arrays for return
    const addBarcodesArray = Array.from(addBarcodes.entries()).map(([barcode, data]) => ({
      barcode: barcode,
      itemName: data.itemName,
      row: data.row
    }));
    
    const dropBarcodesArray = Array.from(dropBarcodes.entries()).map(([barcode, data]) => ({
      barcode: barcode,
      itemName: data.itemName,
      row: data.row
    }));
    
    Logger.log(`Found ${duplicates.length} duplicates, ${addBarcodesArray.length} add barcodes, ${dropBarcodesArray.length} drop barcodes`);
    
    return {
      addBarcodes: addBarcodesArray,
      dropBarcodes: dropBarcodesArray,
      duplicates: duplicates
    };
    
  } catch (error) {
    Logger.log(`❌ Error in checkForDuplicates: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
} 