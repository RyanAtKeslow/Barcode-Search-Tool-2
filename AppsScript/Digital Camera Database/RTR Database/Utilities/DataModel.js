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

 

function FetchUserEmailFromInfo(firstName) {
  // Special case for JT
  if (firstName.trim().toLowerCase() === 'jt') {
    return {
      email: 'jthomas@keslowcamera.com',
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