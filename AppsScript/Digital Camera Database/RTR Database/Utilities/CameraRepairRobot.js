/**
 * Camera Repair Robot - Sends Google Chat notifications when cameras are marked for repair
 */

function cameraRepairRobot(cameraName, cameraSN, cameraBC, userName = null) {
  // Use provided userName or fallback to current session user
  let fullName;
  if (userName) {
    fullName = userName;
  } else {
    const userInfo = fetchUserInfoFromEmail();
    fullName = userInfo.fullName;
  }
  const message = `${cameraName} SN${cameraSN} BC${cameraBC} has been put into REPAIR by ${fullName}`;
  
  // Send message to Google Chat
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAAwjNel5g/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Ihcxnkjwf2Dr9oRTMGo-t0WtRX_l8vPK0NJ9bWXPTX4';
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