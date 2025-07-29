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

function cameraRepairRobot(cameraName, cameraSN, cameraBC) {
  const userInfo = fetchUserInfoFromEmail();
  const message = `${cameraName} SN${cameraSN} BC${cameraBC} has been put into REPAIR by ${userInfo.fullName}`;
  
  // Send message to Google Chat
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAAwjNel5g/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Ihcxnkjwf2Dr9oRTMGo-t0WtRX_l8vPK0NJ9bWXPTX4'; // You'll need to replace this with your actual webhook URL
  const payload = {
    'text': message
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };
  
  try {
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log('Repair message sent successfully');
  } catch (error) {
    Logger.log('Error sending repair message: ' + error.toString());
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