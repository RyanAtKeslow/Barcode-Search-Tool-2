/**
 * Venice 2 Sensor Robot - Sends Google Chat notifications for sensor block changes
 */

function venice2SensorRobot(message) {
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAAwjNel5g/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=-TLPuHM_flh2NmnTQtSpJIyVkL2RtMmxZajHEmX51Ew';
  const payload = JSON.stringify({ text: message });
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload
  };
  UrlFetchApp.fetch(webhookUrl, options);
} 