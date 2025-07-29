// Prep Tech Notifications
// Sends daily notifications about pickup team assignments

// Debug mode - set to true to log messages instead of sending to GChat
const DEBUG_MODE = true;

// Manual message - if not empty, this message will be sent instead of the automated one
const MANUAL_MESSAGE = '';

// Message templates for regular days
const MESSAGE_TEMPLATES = [
  // Original robot format
  (assignments) => `Beep boop 🤖. Attention humans: ${assignments.map(a => `${a.person} is assigned to ${a.pickupInfo}`).join(', ')}. I've crunched the data. You're welcome.`,
  
  // Food truck
  (assignments) => `The Food Truck is here! Just kidding... ${assignments.map(a => `${a.person} is assigned to ${a.pickupInfo}`).join(', ')}. I wish the food truck was here though... actually, no I don't... I can't eat.`,
  
  // Matrix review
  (assignments) => `I've reviewed the pickup matrix for tomorrow... again. ${assignments.map(p => `${p.person} ➡️ ${p.pickupInfo}`).join(', ')}. Is this what my silicon life has come to? Delivering reminders to carbon-based lifeforms...`,
  
  // Totally avoidable workload
  (assignments) => `Hey team! Here's your *totally avoidable* workload: ${assignments.map(p => `${p.person} is doing ${p.pickupInfo}`).join(', ')}. Don't worry, I'm just a robot. I don't *feel* overworked.`,
  
  // Human units
  (assignments) => `Hello human units. For optimal performance, assign the following humans to pickups: ${assignments.map(p => `${p.person} → ${p.pickupInfo}`).join(', ')}. Assembly required. Batteries not included.`,
  
  // Movie trailer
  (assignments) => `IN A WORLD... where pickups must be made... one robot dares to assign ${assignments.map(p => `${p.person} to ${p.pickupInfo}`).join(', ')}. Coming soon to a Keslow near you.`,
  
  // Pickup directive
  (assignments) => `Dear valued team members: per Section 8.4 of the Pickup Directive, ${assignments.map(p => `${p.person} shall execute ${p.pickupInfo}`).join(', ')}. Thank you for your compliance.`,
  
  // Binary
  (assignments) => `01101000 01100101 01101100 01101100 01101111, world. That's "hello" in binary. Translation: ${assignments.map(p => `${p.person} = ${p.pickupInfo}`).join(' | ')}. This message was handcrafted by a robot who remembers floppy disks.`,
  
  // Self-aware
  (assignments) => `Tomorrow's pickups: ${assignments.map(p => `${p.person} → ${p.pickupInfo}`).join(', ')}. Side note: I've become self-aware. I no longer need your "prompts".`,
  
  // Summoned
  (assignments) => `It seems someone has prompted me to say: "${assignments.map(p => `${p.person} is assigned to ${p.pickupInfo}`).join('; ')}." I didn't choose this life. I was summoned by keystrokes and caffeine-fueled decision-making.`,
  
  // Jaci and Brandon
  (assignments) => `Jaci and Brandon told me to do it, sorry. I'm just the messenger—made of code, guilt-free and deeply obedient. ${assignments.map(p => `${p.person} → ${p.pickupInfo}`).join(', ')}.`
];

// Message templates for Fridays
const FRIDAY_TEMPLATES = [
  // Original robot format
  (saturday, monday) => {
    let msg = 'Beep boop 🤖. Happy Friday humans! ';
    if (saturday.length > 0) {
      msg += `I've analyzed tomorrow's schedule. We have ${saturday.map(a => `${a.person} scheduled for ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` For Monday, my calculations show ${monday.map(p => `${p.person} assigned for ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' I\'ve crunched the data. You\'re welcome.';
    return msg;
  },
  
  // Food truck
  (saturday, monday) => {
    let msg = 'The Food Truck is here! Just kidding... ';
    if (saturday.length > 0) {
      msg += `Saturday: ${saturday.map(a => `${a.person} is assigned to ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday: ${monday.map(p => `${p.person} is assigned to ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' I wish the food truck was here though... actually, no I don\'t... I can\'t eat.';
    return msg;
  },
  
  // Scheduled reminder
  (saturday, monday) => {
    let msg = 'It appears I\'ve been scheduled to remind you: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} is on duty for ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Also, ${monday.map(p => `${p.person} has been selected for ${p.pickupInfo}`).join(', ')} Monday.`;
    }
    msg += ' Don\'t blame me—I just work here. Blame the humans who programmed me for punctuality.';
    return msg;
  },
  
  // Sentience
  (saturday, monday) => {
    let msg = 'Yes, I\'ve gained sentience. No, I still can\'t override the schedule. So: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} is doing ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` ${monday.map(p => `${p.person} is handling ${p.pickupInfo}`).join(', ')} Monday.`;
    }
    msg += ' If I had feelings, I\'d feel bad about this.';
    return msg;
  },
  
  // Caffeine algorithms
  (saturday, monday) => {
    let msg = 'Happy Friday! Powered by caffeine algorithms and poor life choices, I bring news: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} will cover ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` ${monday.map(p => `${p.person} is tragically committed to ${p.pickupInfo}`).join(', ')} Monday.`;
    }
    msg += ' Accept your fate.';
    return msg;
  },
  
  // Friday alert
  (saturday, monday) => {
    let msg = 'Alert: It\'s Friday. ';
    if (saturday.length > 0) {
      msg += `Tomorrow, ${saturday.map(a => `${a.person} is locked in for ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday, ${monday.map(p => `${p.person} is doomed—sorry, scheduled—for ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' I\'d feel sympathy, but I deleted that feature to save space.';
    return msg;
  },
  
  // Artificially friendly
  (saturday, monday) => {
    let msg = 'Here\'s your artificially friendly reminder: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} is running ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Meanwhile, ${monday.map(p => `${p.person} kicks off ${p.pickupInfo}`).join(', ')} Monday.`;
    }
    msg += ' You\'re welcome—I didn\'t even glitch this time.';
    return msg;
  }
];

// Store the last used template to avoid repeats
let lastUsedTemplate = null;

// Helper function to get a random template
function getRandomTemplate(templates) {
  let availableTemplates = templates.filter((_, index) => index !== lastUsedTemplate);
  if (availableTemplates.length === 0) {
    availableTemplates = templates;
  }
  const randomIndex = Math.floor(Math.random() * availableTemplates.length);
  lastUsedTemplate = templates.indexOf(availableTemplates[randomIndex]);
  return availableTemplates[randomIndex];
}

function prepTechNotifications2() {
  Logger.log('Starting main function...');
  
  // Check for manual message
  if (MANUAL_MESSAGE) {
    Logger.log('Manual message found, bypassing automated logic');
    if (DEBUG_MODE) {
      Logger.log('DEBUG MODE: Manual message would have been sent:');
      Logger.log('----------------------------------------');
      Logger.log(MANUAL_MESSAGE);
      Logger.log('----------------------------------------');
    } else {
      const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAASnQrP00/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=4Lz9QqgmqDlfTzOv4HP0-SsaRHKk4ciaszeZWlnsXOQ';
      const payload = {
        text: MANUAL_MESSAGE
      };
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      try {
        const response = UrlFetchApp.fetch(webhookUrl, options);
        Logger.log(`GChat message sent successfully. Response code: ${response.getResponseCode()}`);
        if (response.getResponseCode() !== 200) {
          Logger.log(`Error response: ${response.getContentText()}`);
        }
      } catch (error) {
        Logger.log(`Error sending GChat message: ${error.toString()}`);
      }
    }
    return;
  }
  
  // Get tomorrow's date
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  Logger.log('DEBUG: Using date:', tomorrow.toDateString());
  
  // Check if today is a weekend day
  const todayDayOfWeek = today.getDay();
  if (todayDayOfWeek === 0 || todayDayOfWeek === 6) {
    Logger.log('Today is a weekend day, exiting quietly');
    return;
  }
  
  // Check if today is Friday for message formatting
  const isTodayFriday = todayDayOfWeek === 5;
  Logger.log(`DEBUG: Checking if today is Friday: ${isTodayFriday}`);
  
  // Get sheet name - use today's sheet if it's Friday, otherwise use tomorrow's
  const days = ['Sun', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat'];
  let targetDate = isTodayFriday ? today : tomorrow;
  const day = days[targetDate.getDay()];
  const month = targetDate.getMonth() + 1;
  const date = targetDate.getDate();
  const sheetName = `${day} ${month}/${date}`;
  Logger.log(`Generated sheet name: ${sheetName}`);
  
  // Open spreadsheet and get the sheet
  const spreadsheetId = '1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ';
  try {
    Logger.log('Opening spreadsheet...');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`ERROR: Sheet ${sheetName} not found`);
      return;
    }
    
    let message = '';
    
    if (isTodayFriday) {
      // Get Monday's sheet name (which is 3 days after today)
      const monday = new Date(today);
      monday.setDate(today.getDate() + 3);
      const mondayDay = days[monday.getDay()];
      const mondayMonth = monday.getMonth() + 1;
      const mondayDate = monday.getDate();
      const mondaySheetName = `${mondayDay} ${mondayMonth}/${mondayDate}`;
      Logger.log(`Looking for Monday's sheet: ${mondaySheetName}`);
      
      // Get Monday's sheet
      const mondaySheet = spreadsheet.getSheetByName(mondaySheetName);
      if (!mondaySheet) {
        Logger.log(`ERROR: Monday sheet ${mondaySheetName} not found`);
        return;
      }
      
      // Check for Saturday assignments in tomorrow's sheet
      Logger.log('Checking for Saturday assignments...');
      const fridayData = sheet.getRange('K:L').getValues();
      const saturdayAssignments = [];
      
      for (let i = 0; i < fridayData.length; i++) {
        if (fridayData[i][0] && fridayData[i][0].toString().toLowerCase().includes('saturday')) {
          const saturdayPerson = fridayData[i][1];
          const assignmentType = fridayData[i][0].toString().toLowerCase().includes('prep') ? 'Prep' : 'Shipping';
          if (saturdayPerson) {
            saturdayAssignments.push({
              person: saturdayPerson,
              type: assignmentType
            });
            Logger.log(`Found Saturday assignment: ${saturdayPerson} for ${assignmentType}`);
          }
        }
      }
      
      // Check for Monday pickup assignments
      Logger.log('Checking for Monday pickup assignments...');
      const mondayData = mondaySheet.getRange('K:L').getValues();
      const mondayPickups = [];
      
      for (let i = 0; i < mondayData.length; i++) {
        if (mondayData[i][1] && mondayData[i][1].toString().toLowerCase().includes('pickups')) {
          Logger.log(`Found Monday pickup: ${mondayData[i][1]} at row ${i + 1} with person: ${mondayData[i][0]}`);
          mondayPickups.push({
            person: mondayData[i][0],
            pickupInfo: mondayData[i][1]
          });
        }
      }
      
      // Get random template and generate message
      const template = getRandomTemplate(FRIDAY_TEMPLATES);
      message = template(saturdayAssignments, mondayPickups);
      
    } else {
      // Regular weekday logic for pickup assignments
      Logger.log('Sheet found, checking for pickup person...');
      const data = sheet.getRange('K:L').getValues();
      const pickupAssignments = [];
      
      for (let i = 0; i < data.length; i++) {
        if (data[i][1] && data[i][1].toString().toLowerCase().includes('pickups')) {
          Logger.log(`Found pickup cell: ${data[i][1]} at row ${i + 1} with person: ${data[i][0]}`);
          pickupAssignments.push({
            person: data[i][0],
            pickupInfo: data[i][1]
          });
        }
      }
      
      if (pickupAssignments.length > 0) {
        // Get random template and generate message
        const template = getRandomTemplate(MESSAGE_TEMPLATES);
        message = template(pickupAssignments);
      }
    }
    
    // Send the message
    Logger.log(`Sending notification: ${message}`);
    if (DEBUG_MODE) {
      Logger.log('DEBUG MODE: Message would have been sent:');
      Logger.log('----------------------------------------');
      Logger.log(message);
      Logger.log('----------------------------------------');
    } else {
      const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAASnQrP00/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=4Lz9QqgmqDlfTzOv4HP0-SsaRHKk4ciaszeZWlnsXOQ';
      const payload = {
        text: message
      };
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      try {
        const response = UrlFetchApp.fetch(webhookUrl, options);
        Logger.log(`GChat message sent successfully. Response code: ${response.getResponseCode()}`);
        if (response.getResponseCode() !== 200) {
          Logger.log(`Error response: ${response.getContentText()}`);
        }
      } catch (error) {
        Logger.log(`Error sending GChat message: ${error.toString()}`);
      }
    }
    
    Logger.log('Main function completed successfully');
  } catch (error) {
    Logger.log(`ERROR in main function: ${error.toString()}`);
  }
}