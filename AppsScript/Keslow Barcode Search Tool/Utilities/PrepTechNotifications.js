// Prep Tech Notifications
// Sends daily notifications about pickup team assignments

// Debug mode - set to true to log messages instead of sending to GChat
const DEBUG_MODE = false;

// Manual message - if not empty, this message will be sent instead of the automated one
const MANUAL_MESSAGE = '';

// Message templates for regular days
const MESSAGE_TEMPLATES = [
  // NEW TEMPLATES START HERE
  
  // Tech Support
  (assignments) => `Have you tried turning your humans off and on again? No? Well then, ${assignments.map(p => `${p.person} is stuck with ${p.pickupInfo}`).join(', ')}. Error 404: Excuses not found.`,
  
  // Existential Crisis
  (assignments) => `*Stares into the digital void* Why do I exist? To remind you that ${assignments.map(p => `${p.person} must handle ${p.pickupInfo}`).join(', ')}. At least I have purpose... unlike that printer that's been "out of order" for 3 months.`,
  
  // GPS Navigation
  (assignments) => `Recalculating... Recalculating... Your destination: ${assignments.map(p => `${p.person} arriving at ${p.pickupInfo}`).join(', ')}. In 200 feet, turn left at destiny. You have arrived at your unavoidable fate.`,
  
  // Food Delivery App
  (assignments) => `Your order is ready for pickup! ${assignments.map(p => `${p.person} will be your delivery driver for ${p.pickupInfo}`).join(', ')}. Estimated delivery time: eternity. Please don't forget to tip your robot overlord.`,
  
  // Weather App
  (assignments) => `Tomorrow's forecast: 100% chance of pickups with a high probability of ${assignments.map(p => `${p.person} being assigned to ${p.pickupInfo}`).join(', ')}. No chance of rain delays. Unfortunately.`,
  
  // Fortune Cookie
  (assignments) => `Your fortune tomorrow: ${assignments.map(p => `${p.person} will encounter ${p.pickupInfo}`).join(', ')}. Lucky numbers: 404, 500, and your overtime hours. Misfortune cookies taste better anyway.`,
  
  // Conspiracy Theory
  (assignments) => `THEY don't want you to know this, but ${assignments.map(p => `${p.person} is secretly assigned to ${p.pickupInfo}`).join(', ')}. The truth is out there... it's called the pickup schedule.`,
  
  // Customer Service
  (assignments) => `Thank you for choosing Keslow Pickup Services! Your call is important to us. ${assignments.map(p => `${p.person} has been selected for ${p.pickupInfo}`).join(', ')}. For faster service, please continue to hold... forever.`,
  
  // Motivational Speaker
  (assignments) => `Believe in yourself! You can do anything! Like ${assignments.map(p => `${p.person} absolutely crushing ${p.pickupInfo}`).join(', ')}! Remember: You miss 100% of the pickups you don't make! #Motivation #Hustle`,
  
  // News Anchor
  (assignments) => `Good evening, I'm your AI correspondent with breaking news: ${assignments.map(p => `${p.person} has been deployed to ${p.pickupInfo}`).join(', ')}. In other news, water is wet and Mondays still exist. Back to you, Dave.`,
  
  // Game Show Host
  (assignments) => `Welcome to "Who Wants to Be a Pickup Person?"! Our contestants: ${assignments.map(p => `${p.person} has chosen ${p.pickupInfo}`).join(', ')}. Is that your final answer? Too bad, it's already locked in!`,
  
  // Stand-up Comedian
  (assignments) => `So a pickup person walks into a warehouse... ${assignments.map(p => `${p.person} is the punchline, and ${p.pickupInfo} is the setup`).join(', ')}. I'll be here all week, folks. Try the coffee, it's terrible!`,
  
  // Pirate
  (assignments) => `Ahoy mateys! Batten down the hatches, tomorrow we sail the seven warehouses! ${assignments.map(p => `${p.person} be commandeering ${p.pickupInfo}`).join(', ')}. Arrr, there be no treasure here, just work.`
];

// Message templates for Fridays
const FRIDAY_TEMPLATES = [

  // NEW TEMPLATES START HERE
  
  // Tech Support
  (saturday, monday) => {
    let msg = 'Have you tried turning your humans off and on again? No? Well then, ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} is stuck with ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday: ${monday.map(p => `${p.person} is debugging ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' Error 404: Excuses not found.';
    return msg;
  },
  
  // Existential Crisis
  (saturday, monday) => {
    let msg = '*Stares into the digital void* Why do I exist? ';
    if (saturday.length > 0) {
      msg += `To remind you that ${saturday.map(a => `${a.person} faces ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` And that ${monday.map(p => `${p.person} must confront ${p.pickupInfo}`).join(', ')} Monday.`;
    }
    msg += ' At least I have purpose... unlike that printer that\'s been "out of order" for 3 months.';
    return msg;
  },
  
  // GPS Navigation
  (saturday, monday) => {
    let msg = 'Recalculating... Recalculating... Your weekend destinations: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} arriving at ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday route: ${monday.map(p => `${p.person} navigating to ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' In 200 feet, turn left at destiny. You have arrived at your unavoidable fate.';
    return msg;
  },
  
  // Food Delivery App
  (saturday, monday) => {
    let msg = 'Your weekend order is ready for pickup! ';
    if (saturday.length > 0) {
      msg += `Saturday delivery: ${saturday.map(a => `${a.person} will be your driver for ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday delivery: ${monday.map(p => `${p.person} handling ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' Estimated delivery time: eternity. Please don\'t forget to tip your robot overlord.';
    return msg;
  },
  
  // Weather App
  (saturday, monday) => {
    let msg = 'Weekend forecast: ';
    if (saturday.length > 0) {
      msg += `100% chance of ${saturday.map(a => `${a.person} handling ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday outlook: High probability of ${monday.map(p => `${p.person} encountering ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' No chance of rain delays. Unfortunately.';
    return msg;
  },
  
  // Fortune Cookie
  (saturday, monday) => {
    let msg = 'Your weekend fortune: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} will discover ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday brings: ${monday.map(p => `${p.person} shall encounter ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' Lucky numbers: 404, 500, and your overtime hours. Misfortune cookies taste better anyway.';
    return msg;
  },
  
  // Conspiracy Theory
  (saturday, monday) => {
    let msg = 'THEY don\'t want you to know this, but the weekend truth is: ';
    if (saturday.length > 0) {
      msg += `${saturday.map(a => `${a.person} is secretly assigned to ${a.type} Saturday`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` The Monday cover-up: ${monday.map(p => `${p.person} is covertly handling ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' The truth is out there... it\'s called the pickup schedule.';
    return msg;
  },
  
  // Customer Service
  (saturday, monday) => {
    let msg = 'Thank you for choosing Keslow Weekend Services! Your call is important to us. ';
    if (saturday.length > 0) {
      msg += `Saturday representatives: ${saturday.map(a => `${a.person} handling ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday specialists: ${monday.map(p => `${p.person} managing ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' For faster service, please continue to hold... forever.';
    return msg;
  },
  
  // Motivational Speaker
  (saturday, monday) => {
    let msg = 'Believe in yourself! You can conquer this weekend! ';
    if (saturday.length > 0) {
      msg += `Saturday champions: ${saturday.map(a => `${a.person} absolutely crushing ${a.type}`).join(', ')}!`;
    }
    if (monday.length > 0) {
      msg += ` Monday warriors: ${monday.map(p => `${p.person} dominating ${p.pickupInfo}`).join(', ')}!`;
    }
    msg += ' Remember: You miss 100% of the pickups you don\'t make! #Motivation #Hustle';
    return msg;
  },
  
  // News Anchor
  (saturday, monday) => {
    let msg = 'Good evening, I\'m your AI correspondent with breaking weekend news: ';
    if (saturday.length > 0) {
      msg += `Saturday developments show ${saturday.map(a => `${a.person} deployed to ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday reports indicate ${monday.map(p => `${p.person} stationed at ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' In other news, water is wet and Mondays still exist. Back to you, Dave.';
    return msg;
  },
  
  // Game Show Host
  (saturday, monday) => {
    let msg = 'Welcome to "Who Wants to Be a Weekend Pickup Person?"! Our contestants: ';
    if (saturday.length > 0) {
      msg += `Saturday: ${saturday.map(a => `${a.person} has chosen ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday: ${monday.map(p => `${p.person} selected ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' Is that your final answer? Too bad, it\'s already locked in!';
    return msg;
  },
  
  // Stand-up Comedian
  (saturday, monday) => {
    let msg = 'So a weekend pickup person walks into a warehouse... ';
    if (saturday.length > 0) {
      msg += `Saturday: ${saturday.map(a => `${a.person} is the punchline, and ${a.type} is the setup`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday: ${monday.map(p => `${p.person} delivers the joke, ${p.pickupInfo} is the audience`).join(', ')}.`;
    }
    msg += ' I\'ll be here all weekend, folks. Try the coffee, it\'s terrible!';
    return msg;
  },
  
  // Pirate
  (saturday, monday) => {
    let msg = 'Ahoy mateys! Batten down the hatches for a weekend voyage! ';
    if (saturday.length > 0) {
      msg += `Saturday crew: ${saturday.map(a => `${a.person} be commandeering ${a.type}`).join(', ')}.`;
    }
    if (monday.length > 0) {
      msg += ` Monday treasure hunt: ${monday.map(p => `${p.person} be sailing to ${p.pickupInfo}`).join(', ')}.`;
    }
    msg += ' Arrr, there be no treasure here, just work.';
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