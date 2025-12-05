/**
 * Lens Mount Robot - Sends Google Chat notifications for lens mount changes on any camera
 */

function lensMountRobot(cameraBodyName, serialNumber, barcodeNumber, oldMountValue, newMountValue, userName) {
  const message = `${cameraBodyName} SN(${serialNumber}) BC(${barcodeNumber}) has been changed from ${oldMountValue} -> ${newMountValue} by ${userName}`;
  
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAAwjNel5g/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=21igE1u4jF3E4TtaoM-dMgiLL7IdQvO_F_opVXrYkbM';
  const payload = JSON.stringify({ text: message });
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload
  };
  UrlFetchApp.fetch(webhookUrl, options);
}

// Backward compatibility alias - accepts old message format
function venice2LensMountRobot(message) {
  // For backward compatibility, if called with a single message string, send it directly
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAAwjNel5g/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=21igE1u4jF3E4TtaoM-dMgiLL7IdQvO_F_opVXrYkbM';
  const payload = JSON.stringify({ text: message });
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload
  };
  UrlFetchApp.fetch(webhookUrl, options);
} 