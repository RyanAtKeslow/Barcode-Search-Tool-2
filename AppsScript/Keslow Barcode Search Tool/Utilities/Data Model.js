/**
 * Data Model - Barcode Dictionary Utilities
 * 
 * This script provides utility functions for managing the Barcode Dictionary
 * and user information within the Keslow Barcode Search Tool system.
 * 
 * Main Functions:
 * 1. backupBarcodeDictionary() - Creates timestamped backups of barcode data
 * 2. checkBarcodeIntegrity() - Validates data integrity between raw, formatted, and dictionary data
 * 3. fetchUserInfoFromEmail() - Retrieves user details from company email list
 * 4. findUsernameInRow2() - Locates user names in sheet row 2 with fallback logic
 * 5. setSelectedMatch() - Handles multiple username matches with user selection dialog
 * 
 * User Management:
 * - Fetches user info from US/CAN company email lists
 * - Handles special cases for specific users (JT, Andy, Vinnie, Gigi, Jaz)
 * - Provides nickname extraction and email lookup functionality
 * 
 * Data Integrity:
 * - Compares barcode counts between raw data, formatted data, and dictionary
 * - Identifies missing barcodes and new barcodes
 * - Ensures data consistency across processing stages
 */
function backupBarcodeDictionary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Barcode Dictionary");
  const data = sheet.getDataRange().getValues();

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm");
  const backupName = `Barcode Dictionary Backup – ${timestamp}`;

  // Check if a folder named "Barcode Backups" exists, otherwise create it
  const folders = DriveApp.getFoldersByName("Barcode Backups");
  const backupFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Barcode Backups");

  // Create new sheet and write data
  const backupFile = SpreadsheetApp.create(backupName);
  const backupSheet = backupFile.getSheets()[0];
  backupSheet.clear();
  backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Move file into the folder
  const file = DriveApp.getFileById(backupFile.getId());
  backupFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // remove from root

  Logger.log("✅ Backup created at: " + backupName);
}

function checkBarcodeIntegrity(rawSheet, formattedSheet, barcodeDictSheet) {
  if (!barcodeDictSheet) {
    throw new Error("❌ 'Barcode Dictionary' sheet not found. Aborting integrity check.");
  }
  const rawData = rawSheet.getDataRange().getValues();
  const formattedData = formattedSheet.getDataRange().getValues();
  const barcodeDictData = barcodeDictSheet.getRange(2, 7, barcodeDictSheet.getLastRow() - 1).getValues(); // Col G

  // OLD BARCODES (from Barcode Dictionary)
  const oldBarcodes = new Set();
  for (let row of barcodeDictData) {
    const barcodes = String(row[0]).split("|").map(b => b.trim());
    barcodes.forEach(b => oldBarcodes.add(b));
  }

  // RAW BARCODES (from dump before formatting)
  const rawBarcodes = new Set();
  const headerMap = {
    UUID: 0,
    Equipment: 1,
    Barcode: 2,
    Category: 3,
    Status: 4,
    Owner: 5,
    Location: 6
  };

  for (let i = 1; i < rawData.length; i++) {
    const barcode = String(rawData[i][headerMap.Barcode]).trim();
    if (barcode) rawBarcodes.add(barcode);
  }

  // FORMATTED BARCODES (after grouping)
  const formattedBarcodes = new Set();
  for (let i = 1; i < formattedData.length; i++) {
    const cell = formattedData[i][6]; // Column G (Barcodes)
    if (typeof cell === 'string') {
      const barcodes = cell.split("|").map(b => b.trim());
      barcodes.forEach(b => formattedBarcodes.add(b));
    }
  }

  // Compare sets
  const missingBarcodes = [...rawBarcodes].filter(b => !formattedBarcodes.has(b));
  const newBarcodes = [...rawBarcodes].filter(b => !oldBarcodes.has(b));
  const integrityPassed = missingBarcodes.length === 0;

  return {
    integrityPassed,
    missingBarcodes,
    newBarcodes
  };
}

function fetchUserInfoFromEmail() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById('1NVDIITUWr52e-sTmz-YJWY9sPOunx1VGIWQYNPqCKg8');
  const sheetsToCheck = ['US', 'CAN'];
  let userInfo = {
    city: 'City not found',
    fullName: 'user not found',
    firstName: 'no first name',
    lastName: 'no last name',
    email: 'user email not found'
  };

  for (let sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); // A-E
    for (let row of data) {
      if (row[4] && row[4].toString().toLowerCase() === email.toLowerCase()) {
        userInfo = {
          city: row[0],
          fullName: row[1],
          firstName: row[2],
          lastName: row[3],
          email: row[4]
        };
        Logger.log(`User found in ${sheetName}: ${JSON.stringify(userInfo)}`);
        return userInfo;
      }
    }
  }
  Logger.log('User not found in Company Email List.');
  return userInfo;
}

function FetchUserFromFormSubmitViaEmail(email) {
  const ss = SpreadsheetApp.openById('1NVDIITUWr52e-sTmz-YJWY9sPOunx1VGIWQYNPqCKg8');
  const sheetsToCheck = ['US', 'CAN'];
  let userInfo = {
    city: 'City not found',
    fullName: 'user not found',
    firstName: 'no first name',
    lastName: 'no last name',
    email: 'user email not found'
  };

  for (let sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); // A-E
    for (let row of data) {
      if (row[4] && row[4].toString().toLowerCase() === email.toLowerCase()) {
        userInfo = {
          city: row[0],
          fullName: row[1],
          firstName: row[2],
          lastName: row[3],
          email: row[4]
        };
        Logger.log(`User found in ${sheetName}: ${JSON.stringify(userInfo)}`);
        return userInfo;
      }
    }
  }
  Logger.log('User not found in Company Email List.');
  return userInfo;
}

/**
 * Handles the selection of a username cell when multiple matches are found
 * @param {Array} usernameMatches - Array of objects containing cellA1 and value for each match
 * @returns {Object} The selected match object containing cellA1 and value
 */
function setSelectedMatch(usernameMatches) {
  try {
    Logger.log(`Handling ${usernameMatches.length} username matches`);
    
    // Create and show the selection dialog
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
          const selectedMatch = ${JSON.stringify(usernameMatches)}[index];
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .withFailureHandler((error) => alert(error))
            .handleSelectedMatch(selectedMatch);
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(200);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Correct Entry');
    
    // Wait for the selection to be made and stored
    let selectedMatch = null;
    let attempts = 0;
    const maxAttempts = 30; // 30 seconds timeout
    
    while (!selectedMatch && attempts < maxAttempts) {
      const storedMatch = PropertiesService.getScriptProperties().getProperty('selectedMatch');
      if (storedMatch) {
        selectedMatch = JSON.parse(storedMatch);
        PropertiesService.getScriptProperties().deleteProperty('selectedMatch');
        break;
      }
      Utilities.sleep(1000);
      attempts++;
    }
    
    if (!selectedMatch) {
      throw new Error('No selection was made within the timeout period');
    }
    
    return selectedMatch;
  } catch (error) {
    Logger.log(`❌ Error in setSelectedMatch: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Handles the selected match from the dialog
 * @param {Object} match - The selected match object
 */
function handleSelectedMatch(match) {
  try {
    Logger.log(`Handling selected match: ${JSON.stringify(match)}`);
    PropertiesService.getScriptProperties().setProperty('selectedMatch', JSON.stringify(match));
  } catch (error) {
    Logger.log(`❌ Error in handleSelectedMatch: ${error.toString()}`);
    throw error;
  }
}

/**
 * Fetches the current user's email and nickname
 * @returns {Object} Object containing email and nickname
 */
function fetchUserEmailandNickname() {
  try {
    const email = Session.getActiveUser().getEmail();
    const userInfo = FetchUserFromFormSubmitViaEmail(email);
    
    // Extract nickname (first name) from user info
    const nickname = userInfo.firstName || email.split('@')[0];
    
    Logger.log(`Fetched user info - Email: ${email}, Nickname: ${nickname}`);
    
    return {
      email: email,
      nickname: nickname
    };
  } catch (error) {
    Logger.log(`❌ Error in fetchUserEmailandNickname: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Searches for user's name in row 2 using a two-step process:
 * 1. First tries to find first/last name from company email list
 * 2. If no match found, falls back to email nickname
 * @param {Sheet} sheet - The sheet to search in
 * @returns {Object} Object containing username cell and value
 */
function findUsernameInRow2(sheet) {
  try {
    Logger.log('=== findUsernameInRow2 START ===');
    const userInfo = fetchUserInfoFromEmail();
    const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let usernameMatches = [];
    
    // First attempt: Search using first name, last name, or combination
    const searchTerms = [
      userInfo.firstName,
      userInfo.lastName,
      `${userInfo.firstName} ${userInfo.lastName}`,
      `${userInfo.lastName} ${userInfo.firstName}`
    ].filter(term => term && term !== 'no first name' && term !== 'no last name');
    
    Logger.log(`Searching with terms: ${JSON.stringify(searchTerms)}`);
    
    for (let j = 0; j < row2.length; j++) {
      const cellValue = row2[j];
      if (!cellValue) continue;
      
      const cellValueLower = cellValue.toString().toLowerCase();
      
      // Check if any search term matches
      for (const term of searchTerms) {
        if (cellValueLower.includes(term.toLowerCase())) {
          // Only add if this is a new match
          const isDuplicate = usernameMatches.some(match => 
            match.value.toLowerCase() === cellValue.toString().toLowerCase()
          );
          
          if (!isDuplicate) {
            usernameMatches.push({
              cellA1: sheet.getRange(2, j + 1).getA1Notation(),
              value: cellValue.toString()
            });
            Logger.log(`Found match with term "${term}": ${cellValue}`);
          }
        }
      }
    }
    
    // If no matches found, try with email nickname
    if (usernameMatches.length === 0) {
      Logger.log('No matches found with name, trying email nickname...');
      const { nickname } = fetchUserEmailandNickname();
      
      for (let j = 0; j < row2.length; j++) {
        const cellValue = row2[j];
        if (cellValue && cellValue.toString().toLowerCase().includes(nickname.toLowerCase())) {
          usernameMatches.push({
            cellA1: sheet.getRange(2, j + 1).getA1Notation(),
            value: cellValue.toString()
          });
          Logger.log(`Found match with nickname "${nickname}": ${cellValue}`);
        }
      }
    }
    
    Logger.log(`Found ${usernameMatches.length} total matches`);
    Logger.log('=== findUsernameInRow2 END ===');
    
    return usernameMatches;
  } catch (error) {
    Logger.log(`❌ Error in findUsernameInRow2: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

function FetchUserEmailFromInfo(firstName) {
  // Special case for JT
  if (firstName.trim().toLowerCase() === 'jt') {
    return {
      email: 'jthomas@keslowcamera.com',
      city: 'Los Angeles'
    };
  }

  // Special case for Andy
  if (firstName.trim().toLowerCase() === 'andy') {
    return {
      email: 'arudolph@keslowcamera.com',
      city: 'Los Angeles'
    };
  }

  // Special case for Vinnie
  if (firstName.trim().toLowerCase() === 'vinnie') {
    return {
      email: 'vdemaria@keslowcamera.com',
      city: 'Los Angeles'
    };
  }

  // Special case for Gigi
  if (firstName.trim().toLowerCase() === 'gigi') {
    return {
      email: 'giovannavittone@keslowcamera.com',
      city: 'Los Angeles'
    };
  }

  // Special case for Jaz
  if (firstName.trim().toLowerCase() === 'jaz') {
    return {
      email: 'jasmine@keslowcamera.com',
      city: 'Los Angeles'
    };
  }

  const ss = SpreadsheetApp.openById('1NVDIITUWr52e-sTmz-YJWY9sPOunx1VGIWQYNPqCKg8');
  const sheetsToCheck = ['US', 'CAN'];
  
  for (let sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); // A-E
    for (let row of data) {
      if (row[2] && row[2].toString().toLowerCase() === firstName.trim().toLowerCase()) {
        return {
          email: row[4],
          city: row[0]
        };
      }
    }
  }
  
  Logger.log(`No email found for first name: ${firstName}`);
  return {
    email: null,
    city: null
  };
} 