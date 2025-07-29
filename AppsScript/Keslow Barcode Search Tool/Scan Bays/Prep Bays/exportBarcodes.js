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

function prepBayAddExport() {
  try {
    Logger.log('Starting prepBayAddExport...');
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
    
    // Get the add-barcodes column (one column left of username)
    const addBarcodesCol = usernameCell.getColumn() - 1;
    Logger.log(`Add barcodes column: ${addBarcodesCol}`);
    
    // Pull barcodes column once; we don’t need the neighbouring item-name column
    const lastRow = sheet.getLastRow();
    const columnA = sheet.getRange(4, addBarcodesCol, lastRow - 3, 1).getValues();
    
    // Scan once from bottom to find both export-tag row and last data row
    let exportTagRow = -1;
    let lastDataRow = -1;
    for (let i = columnA.length - 1; i >= 0; i--) {
      const cellVal = columnA[i][0];
      if (!cellVal) continue;
      const str = cellVal.toString();
      if (lastDataRow === -1 && !str.includes("Above was exported @")) {
        lastDataRow = i + 4; // offset for header rows
      }
      if (exportTagRow === -1 && str.includes("Above was exported @")) {
        exportTagRow = i + 4;
      }
      if (lastDataRow !== -1 && exportTagRow !== -1) break; // found both
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
    const barcodeRange = sheet.getRange(startRow, addBarcodesCol, lastDataRow - startRow + 1, 1);
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
    
    // Create filename with timestamp and prep bay number
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd_HH-mm");
    const cellValue = usernameCell.getValue();
    const prepBayNumberForExport = Math.floor((usernameCell.getColumn() - 3) / 5) + 1;
    Logger.log(`Prep Bay Number: ${prepBayNumberForExport}`);
    const filename = `${cellValue}_${timestamp}_adds_Bay${prepBayNumberForExport}.csv`;
    
    // Create CSV file in Drive
    const file = DriveApp.createFile(filename, csvContent, MimeType.CSV);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Create or get the 'Prep Bay Scans' folder in the user's Drive and create file directly inside it
    let folders = DriveApp.getFoldersByName('Prep Bay Scans');
    let prepBayFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Prep Bay Scans');
    prepBayFolder.createFile(file); // Drive API allows createFile on folder object so file is born inside folder

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
    sheet.getRange(lastDataRow + 1, addBarcodesCol).setValue(exportTag);
    
    Logger.log(`✅ Export complete | items: ${filteredData.length}`);
    
    // Get job name from username cell value
    const jobName = cellValue.toString().trim();
    
    // Send digital cameras to database with Pulled status
    // Pass the filtered barcodes array, username, and job name
    const barcodesArray = Array.from(new Set(filteredData.map(row => row[0])));
    SendStatus("Pulled", barcodesArray, username, jobName, userEmail);
    
  } catch (error) {
    Logger.log(`❌ Error in prepBayAddExport: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Callback after user clicks download link. Re-runs SendStatus to ensure Pulled status is recorded
 * even if the initial call failed (rare), and handles multi-prep-bay username selection.
 */
function afterExportComplete() {
  try {
    const { email: userEmail, nickname: username } = fetchUserEmailandNickname();
    const sheet = SpreadsheetApp.getActiveSheet();
    // Find username cells in row 2
    const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const matches = [];
    row2.forEach((v, idx) => {
      if (v && v.toString().toLowerCase().includes(username.toLowerCase())) {
        matches.push({ cell: sheet.getRange(2, idx + 1), value: v });
      }
    });

    let usernameCell;
    if (matches.length === 0) {
      Logger.log('❌ No username match found in prep bays sheet');
      return;
    } else if (matches.length === 1) {
      usernameCell = matches[0].cell;
    } else {
      // Multiple matches – ask user which one (kept original HTML dialog)
      const html = HtmlService.createHtmlOutput(`
        <form id="matchForm">
          <label>Select the correct entry:</label><br/>
          <select id="matchSelect">${matches.map((m,i)=>`<option value="${i}">${m.value}</option>`).join('')}</select><br/>
          <input type="button" value="Select" onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).setSelectedMatch(document.getElementById('matchSelect').value)" />
        </form>`).setWidth(300).setHeight(150);
      SpreadsheetApp.getUi().showModalDialog(html, 'Select Prep Bay');
      return;
    }

    const jobName = usernameCell.getValue().toString().trim();
    const barcodes = sheet.getRange(4, 1, sheet.getLastRow() - 3, 1)
      .getValues()
      .map(r => r[0])
      .filter(b => b && !b.toString().includes('Above was exported @'));
    const uniqueBarcodes = Array.from(new Set(barcodes));
    SendStatus('Pulled', uniqueBarcodes, username, jobName, userEmail);

  } catch (error) {
    Logger.log(`❌ Error in afterExportComplete: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}