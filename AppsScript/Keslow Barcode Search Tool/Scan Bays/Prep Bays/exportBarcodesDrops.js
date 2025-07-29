function prepBayDropExport() {
  try {
    Logger.log('Starting prepBayDropExport...');
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
    const dropBarcodesCol = usernameCell.getColumn() + 2;
    const dropItemNamesCol = usernameCell.getColumn() + 2;  // Item names are in the same column as barcodes
    Logger.log(`Drop barcodes column: ${dropBarcodesCol}, Item names column: ${dropItemNamesCol}`);
    
    // Get all values in both columns starting from row 4
    const columnD = sheet.getRange(4, dropBarcodesCol, sheet.getLastRow() - 3, 1).getValues();
    const columnE = sheet.getRange(4, dropItemNamesCol, sheet.getLastRow() - 3, 1).getValues();
    
    // Find the last row with data and any existing exportTag
    let lastDataRow = -1;
    let exportTagRow = -1;
    
    // First find the most recent export tag, ignoring empty cells
    for (let i = columnD.length - 1; i >= 0; i--) {
      const value = columnD[i][0];
      if (value) {  // Only check non-empty cells
        if (value.toString().includes("Above was exported @")) {
          exportTagRow = i + 4; // +4 because we started at row 4
          break;
        }
      }
    }
    
    // Now find the last barcode after the export tag, ignoring empty cells
    for (let i = columnD.length - 1; i >= 0; i--) {
      const value = columnD[i][0];
      if (value && !value.toString().includes("Above was exported @")) {
        lastDataRow = i + 4; // +4 because we started at row 4
        break;
      }
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
    
    // Filter out any empty cells
    const filteredData = [];
    for (let i = 0; i < barcodes.length; i++) {
      if (barcodes[i][0]) {  // Only include non-empty cells
        filteredData.push([barcodes[i][0]]);  // Only include the barcode
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
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd_HH-mm");
    const cellValue = usernameCell.getValue();
    // Derive prep bay number from usernameCell's column
    const prepBayNumberForExport = Math.floor((usernameCell.getColumn() - 3) / 5) + 1;
    Logger.log(`Prep Bay Number: ${prepBayNumberForExport}`);

    // Create filename with timestamp and prep bay number
    const filename = `${cellValue}_${timestamp}_drops_Bay${prepBayNumberForExport}.csv`;
    
    // Create CSV file in Drive
    const file = DriveApp.createFile(filename, csvContent, MimeType.CSV);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Create or get the 'Prep Bay Scans' folder in the user's Drive
    let folders = DriveApp.getFoldersByName('Prep Bay Scans');
    let prepBayFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Prep Bay Scans');
    prepBayFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Remove from root

    // Create a download URL
    const downloadUrl = file.getDownloadUrl().replace('export=download', 'export=download&confirm=no_antivirus');
    
    // Show success message with download link
    const html = HtmlService.createHtmlOutput(`
      <p>✅ Successfully exported ${filteredData.length} items!</p>
      <p>The file has been saved to your Google Drive and is ready to download.</p>
      <a href="${downloadUrl}" target="_blank" onclick="google.script.run.afterExportComplete()" style="
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
    
    // Set background color of only the barcode cells to light green
    barcodeRange.setBackground("#b7e1cd");
    
    // Add export tag in the next row after the last barcode
    const exportTag = `Above was exported @ ${new Date().toLocaleString()}`;
    sheet.getRange(lastDataRow + 1, dropBarcodesCol).setValue(exportTag);
    
    Logger.log(`✅ Exported ${filteredData.length} items to ${filename}`);
    
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