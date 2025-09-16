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

/**
 * Determines the bay type based on the calling function's name
 * @returns {string} The bay type: "Generic Bay", "Prep Bay", or "Receiving Bay"
 */
function determineBayType() {
  // Get the stack trace to find the calling function
  const stack = new Error().stack;
  
  // Check for specific function patterns in the stack trace
  if (stack.includes('exportGenericBay') || stack.includes('resetGenericBay') || stack.includes('sortGenericBay') || stack.includes('printGenericBay')) {
    return 'Generic Bay';
  } else if (stack.includes('exportBarcodes') || stack.includes('exportBarcodesDrops') || stack.includes('shipPrepBay') || stack.includes('prepBayDropExport') || stack.includes('ShipPrepBay')) {
    return 'Prep Bay';
  } else if (stack.includes('exportReceivingBay') || stack.includes('resetReceivingBay') || stack.includes('sortReceivingBay') || stack.includes('printReceivingBay') || stack.includes('continueResetShippingBay')) {
    return 'Receiving Bay';
  } else {
    // Default fallback - you can change this if needed
    Logger.log(`⚠️ Could not determine bay type from stack trace, defaulting to Generic Bay`);
    Logger.log(`Stack trace: ${stack}`);
    return 'Generic Bay';
  }
}

/**
 * Generates a consistent timestamp format for all CSV files
 * Uses ISO 8601 format with Los Angeles timezone
 * @returns {string} Timestamp in ISO 8601 format with LA timezone (e.g., "2025-09-16T14:30-07:00")
 */
function getConsistentTimestamp() {
  // Get current time in Los Angeles timezone
  const laTimeZone = "America/Los_Angeles";
  const now = new Date();
  
  // Create a date formatter for Los Angeles timezone
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: laTimeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  
  // Format the date in LA timezone
  const parts = formatter.formatToParts(now);
  const year = parts.find(part => part.type === 'year').value;
  const month = parts.find(part => part.type === 'month').value;
  const day = parts.find(part => part.type === 'day').value;
  const hour = parts.find(part => part.type === 'hour').value;
  const minute = parts.find(part => part.type === 'minute').value;
  
  // Get timezone offset for Los Angeles
  const laDate = new Date(now.toLocaleString("en-US", {timeZone: laTimeZone}));
  const utcDate = new Date(now.toLocaleString("en-US", {timeZone: "UTC"}));
  const offsetMs = laDate.getTime() - utcDate.getTime();
  const offsetHours = Math.floor(offsetMs / (1000 * 60 * 60));
  const offsetMinutes = Math.floor((offsetMs % (1000 * 60 * 60)) / (1000 * 60));
  const offsetSign = offsetHours >= 0 ? '+' : '-';
  const offsetString = `${offsetSign}${Math.abs(offsetHours).toString().padStart(2, '0')}:${offsetMinutes.toString().padStart(2, '0')}`;
  
  // Return ISO 8601 format with LA timezone
  return `${year}-${month}-${day}T${hour}:${minute}${offsetString}`;
}

/**
 * Determines the function type based on the calling function's name
 * @returns {string} The function type: "Export", "Reset", "Sort", "Print", "Ship", "Adds", "Drops"
 */
function determineFunctionType() {
  // Get the stack trace to find the calling function
  const stack = new Error().stack;
  
  // Check for specific function patterns in the stack trace
  if (stack.includes('prepBayAddExport')) {
    return 'Adds';
  } else if (stack.includes('prepBayDropExport')) {
    return 'Drops';
  } else if (stack.includes('exportGenericBay') || stack.includes('exportReceivingBay')) {
    return 'Export';
  } else if (stack.includes('resetGenericBay') || stack.includes('resetReceivingBay') || stack.includes('continueResetShippingBay')) {
    return 'Reset';
  } else if (stack.includes('sortGenericBay') || stack.includes('sortReceivingBay')) {
    return 'Sort';
  } else if (stack.includes('printGenericBay') || stack.includes('printReceivingBay')) {
    return 'Print';
  } else if (stack.includes('shipPrepBay') || stack.includes('ShipPrepBay')) {
    return 'Ship';
  } else {
    // Default fallback
    Logger.log(`⚠️ Could not determine function type from stack trace, defaulting to Export`);
    return 'Export';
  }
}

/**
 * Centralized function to save files to Shared Drive
 * This function is called by all scripts that need to save CSV files to the Shared Drive
 * @param {string} folderName - Name of the folder in Shared Drive (deprecated - will be determined by calling function)
 * @param {string} fileName - Name of the file to create
 * @param {string} fileContent - Content of the file
 */
function saveToSharedDrive(folderName, fileName, fileContent) {
  try {
    // Shared Drive ID for Keslow Camera Barcode Archives
    const SHARED_DRIVE_ID = "0AP-pLTczyY0eUk9PVA";
    
    // Determine bay type based on the calling function's name
    const bayType = determineBayType();
    const functionType = determineFunctionType();
    const targetSubfolder = `LA CSV Archives/${bayType}`;
    
    // Add function type suffix to filename if not already present
    let enhancedFileName = fileName;
    if (!fileName.includes(`_${functionType}`)) {
      const lastDotIndex = fileName.lastIndexOf('.');
      if (lastDotIndex !== -1) {
        const nameWithoutExt = fileName.substring(0, lastDotIndex);
        const extension = fileName.substring(lastDotIndex);
        enhancedFileName = `${nameWithoutExt}_${functionType}${extension}`;
      } else {
        enhancedFileName = `${fileName}_${functionType}`;
      }
    }
    
    // Ensure consistent timestamp format in filename (MM-dd-yy_HH-mm)
    // This will standardize timestamps across all files
    const timestampRegex = /\d{2}-\d{2}-\d{2}_\d{2}-\d{2}|\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}|\d{2}\/\d{2}\/\d{2} \d{2}:\d{2}/;
    if (timestampRegex.test(enhancedFileName)) {
      // Replace any existing timestamp with standardized format
      const currentTimestamp = getConsistentTimestamp();
      enhancedFileName = enhancedFileName.replace(timestampRegex, currentTimestamp);
    }
    
    Logger.log(`🔄 Saving to Shared Drive: ${targetSubfolder}/${enhancedFileName}`);
    Logger.log(`🔍 Using Shared Drive ID: ${SHARED_DRIVE_ID}`);
    Logger.log(`📁 Detected bay type: ${bayType}`);
    Logger.log(`🔧 Detected function type: ${functionType}`);
    
    // First, let's test if we can access the Shared Drive
    try {
      const driveInfo = Drive.Files.get(SHARED_DRIVE_ID, {
        supportsAllDrives: true
      });
      Logger.log(`✅ Shared Drive accessible: ${driveInfo.name}`);
    } catch (driveError) {
      Logger.log(`❌ Cannot access Shared Drive: ${driveError.toString()}`);
      Logger.log(`💡 Please verify the Shared Drive ID and permissions`);
      return; // Exit early if we can't access the Shared Drive
    }
    
    // Find the LA CSV Archives folder
    const csvArchivesFolders = Drive.Files.list({
      q: `'${SHARED_DRIVE_ID}' in parents and name='LA CSV Archives' and mimeType='application/vnd.google-apps.folder'`,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'files(id,name)'
    });
    
    let csvArchivesFolderId;
    if (csvArchivesFolders.files && csvArchivesFolders.files.length > 0) {
      csvArchivesFolderId = csvArchivesFolders.files[0].id;
      Logger.log(`✅ Found LA CSV Archives folder (ID: ${csvArchivesFolderId})`);
    } else {
      Logger.log(`❌ Please create a folder called "LA CSV Archives" manually in the Shared Drive first`);
      return; // Exit if folder doesn't exist
    }
    
    // Find or create the bay-specific subfolder
    const bayFolders = Drive.Files.list({
      q: `'${csvArchivesFolderId}' in parents and name='${bayType}' and mimeType='application/vnd.google-apps.folder'`,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'files(id,name)'
    });
    
    let targetFolderId;
    if (bayFolders.files && bayFolders.files.length > 0) {
      targetFolderId = bayFolders.files[0].id;
      Logger.log(`✅ Found existing bay folder: ${bayType} (ID: ${targetFolderId})`);
    } else {
      Logger.log(`❌ Please create a subfolder called "${bayType}" inside "LA CSV Archives" manually in the Shared Drive first`);
      return; // Exit if bay folder doesn't exist
    }
    
    // Create file in Shared Drive folder using DriveApp (more reliable for Shared Drives)
    const blob = Utilities.newBlob(fileContent, 'text/csv', enhancedFileName);
    const folder = DriveApp.getFolderById(targetFolderId);
    const file = folder.createFile(blob);
    
    Logger.log(`✅ File saved to Shared Drive: ${enhancedFileName} (ID: ${file.getId()})`);
    
  } catch (error) {
    Logger.log(`❌ Error saving to Shared Drive: ${error.toString()}`);
    Logger.log(`Error details: ${JSON.stringify(error)}`);
    // Continue execution even if Shared Drive save fails
  }
} 