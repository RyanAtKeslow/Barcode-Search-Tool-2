const AMLFDatabaseCOLS = {
  CAMERA: 2,            // Column B
  SERIAL: 3,            // Column C
  BARCODE: 4,           // Column D
  STATUS: 5,            // Column E (RTR STATUS)
  SERVICE: 6,           // Column F (Most Recent Service Date)
  LOCATION: 7,          // Column G
  MOUNT: 8,             // Column H (MOUNT TYPE)
  OWNER: 9,             // Column I
  CAGE: 10,             // Column J (CAGE TYPE)
  BATTERY: 11,          // Column K (Battery Plate Type)
  FIRMWARE: 12,         // Column L
  HOURS: 13,            // Column M
  NOTES: 14,            // Column N
  LAST_SERVICED_BY: 15, // Column O
  VISUAL: 16            // Column P
};

const AMLFResponseCOLS = {
  TIMESTAMP: 0,         // Column A - Timestamp
  EMAIL: 1,             // Column B - Email Address
  BARCODE: 2,           // Column C - Barcode
  SERIAL: 3,            // Column D - Serial Number
  MOUNT: 4,             // Column E - What Lens Mount is Attached?
  SERVICE_TYPE: 5,      // Column F - Type of Service
  INITIAL_INSPECTION: 6, // Column G - INITIAL INSPECTION
  VISUAL: 7,            // Column H - Visual Impression
  VISUAL_RATING: 8,     // Column I - Visual Impression - Rating
  FIRMWARE: 9,          // Column J - Verify or Update Firmware. Firmware version?
  HOURS: 10,            // Column K - Camera Operating Hours
  USER_INTERFACE: 11,   // Column L - USER INTERFACE, POWER AND I/O
  CAGE_TYPE: 12,        // Column M - Cage Type
  BATTERY_PLATE: 14,    // Column O - Battery Plate Type
  EVF: 15,              // Column P - ELECTRONIC VIEWFINDER (EVF)
  INTERNAL_REC: 16,     // Column Q - INTERNAL RECORDING
  LENS_INTERFACE: 17,   // Column R - LENS INTERFACE
  STATUS: 18,           // Column S - Status
  NOTES: 19             // Column T - Technician Notes (Optional)
};

function AMLFFormSubmit(e) {
  Logger.log('📹 Starting AMLFFormSubmit');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("NEW Alexa Mini LF");
  const dbSheet = ss.getSheetByName("Alexa Mini LF Body Status");

  if (!formSheet || !dbSheet) {
    Logger.log('❌ Sheet validation failed');
    throw new Error("⚠️ One or both sheets not found. Check sheet names.");
  }

  const formData = e.values;
  Logger.log(`📊 Form data received: Serial=${formData[AMLFResponseCOLS.SERIAL]}, Barcode=${formData[AMLFResponseCOLS.BARCODE]}, Email=${formData[AMLFResponseCOLS.EMAIL]}`);

  // Get user info from email
  const userInfo = FetchUserFromFormSubmitViaEmail(formData[AMLFResponseCOLS.EMAIL]);
  console.log(`📧 User info retrieved: ${JSON.stringify(userInfo)}`);

  // Format the service date as M/D/YYYY
  let serviceDate = "";
  if (formData[AMLFResponseCOLS.TIMESTAMP]) {
    const date = new Date(formData[AMLFResponseCOLS.TIMESTAMP]);
    serviceDate = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    console.log(`Formatted service date from '${formData[AMLFResponseCOLS.TIMESTAMP]}' to '${serviceDate}'`);
  }

  // First try to find by serial number in database (Column C)
  let dbData = dbSheet.getRange(2, AMLFDatabaseCOLS.SERIAL, dbSheet.getLastRow() - 1, 1).getValues(); // Column C
  let dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[AMLFResponseCOLS.SERIAL].toString().trim());

  Logger.log('dbData (serials): ' + JSON.stringify(dbData));
  Logger.log('Form serial: [' + formData[AMLFResponseCOLS.SERIAL] + ']');
  Logger.log('dbRowIndex (serial): ' + dbRowIndex);

  // If serial number not found, try to find by barcode (Column D)
  if (dbRowIndex === -1) {
    console.log(`⚠️ Serial Number '${formData[AMLFResponseCOLS.SERIAL]}' not found in Database column C, trying barcode match...`);
    dbData = dbSheet.getRange(2, AMLFDatabaseCOLS.BARCODE, dbSheet.getLastRow() - 1, 1).getValues(); // Column D
    dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[AMLFResponseCOLS.BARCODE].toString().trim());
    
    if (dbRowIndex === -1) {
      console.log(`❌ Neither Serial Number '${formData[AMLFResponseCOLS.SERIAL]}' nor Barcode '${formData[AMLFResponseCOLS.BARCODE]}' found in Database.`);
      
      // Find the first empty row after row 3
      let lastRow = dbSheet.getLastRow();
      let targetRow = 4; // Start from row 4 (after row 3)
      while (targetRow <= lastRow) {
        const rowData = dbSheet.getRange(targetRow, 1, 1, 1).getValues()[0];
        if (!rowData[0]) { // If first cell is empty
          break;
        }
        targetRow++;
      }
      
      // Prepare the new row data
      const newRow = new Array(17).fill(''); // Initialize array with 17 empty strings
      newRow[1] = "ARRI ALEXA Mini LF Camera Body"; // CAMERA
      newRow[AMLFDatabaseCOLS.SERIAL - 1] = formData[AMLFResponseCOLS.SERIAL].toString(); // Ensure serial is treated as text
      newRow[AMLFDatabaseCOLS.BARCODE - 1] = formData[AMLFResponseCOLS.BARCODE].toString(); // Ensure barcode is treated as text
      newRow[AMLFDatabaseCOLS.SERVICE - 1] = serviceDate;
      newRow[AMLFDatabaseCOLS.LOCATION - 1] = userInfo.city; // Add location from user info
      newRow[AMLFDatabaseCOLS.MOUNT - 1] = formData[AMLFResponseCOLS.MOUNT];
      newRow[AMLFDatabaseCOLS.OWNER - 1] = "Keslow Camera"; // Default owner
      newRow[AMLFDatabaseCOLS.CAGE - 1] = formData[AMLFResponseCOLS.CAGE_TYPE];
      newRow[AMLFDatabaseCOLS.BATTERY - 1] = formData[AMLFResponseCOLS.BATTERY_PLATE];
      newRow[AMLFDatabaseCOLS.FIRMWARE - 1] = formData[AMLFResponseCOLS.FIRMWARE];
      newRow[AMLFDatabaseCOLS.HOURS - 1] = formData[AMLFResponseCOLS.HOURS];
      newRow[AMLFDatabaseCOLS.VISUAL - 1] = formData[AMLFResponseCOLS.VISUAL_RATING];
      newRow[AMLFDatabaseCOLS.LAST_SERVICED_BY - 1] = formData[AMLFResponseCOLS.EMAIL];
      
      // Set status based on form data
      const status = formData[AMLFResponseCOLS.STATUS];
      const normalizedStatus = status ? status.trim().toLowerCase() : "";
      if (normalizedStatus === "ready to rent") {
        newRow[AMLFDatabaseCOLS.STATUS - 1] = "RTR";
      } else if (normalizedStatus === "serviced for order") {
        newRow[AMLFDatabaseCOLS.STATUS - 1] = "Pulled";
      } else if (normalizedStatus === "reserve body") {
        newRow[AMLFDatabaseCOLS.STATUS - 1] = "Reserve";
      } else if (normalizedStatus.includes("other")) {
        newRow[AMLFDatabaseCOLS.STATUS - 1] = "UNKNOWN";
      } else if (status && ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"].includes(status)) {
        newRow[AMLFDatabaseCOLS.STATUS - 1] = status;
      }
      
      // Write the new row
      dbSheet.getRange(targetRow, 1, 1, 17).setValues([newRow]);
      
      // Format serial and barcode columns as text to preserve leading zeros
      dbSheet.getRange(targetRow, AMLFDatabaseCOLS.SERIAL).setNumberFormat('@');
      dbSheet.getRange(targetRow, AMLFDatabaseCOLS.BARCODE).setNumberFormat('@');
      
      console.log(`✅ Added new camera entry at row ${targetRow} with serial: ${formData[AMLFResponseCOLS.SERIAL]}, barcode: ${formData[AMLFResponseCOLS.BARCODE]}`);
      
      // Send email notification
      Logger.log('📧 Sending email notification for new camera entry');
      const emailSubject = "Camera Not Found in Database - Alexa Mini LF";
      const emailBody = `A camera was submitted via service form and could not find a match in the database.\n\n` +
                       `Camera Type: Alexa Mini LF\n` +
                       `Serial Number: ${formData[AMLFResponseCOLS.SERIAL]}\n` +
                       `Barcode: ${formData[AMLFResponseCOLS.BARCODE]}\n` +
                       `Submitted by: ${formData[AMLFResponseCOLS.EMAIL]}\n\n` +
                       `Please confirm the accuracy of this information.\n\n` +
                       `The camera has been added to the database at row ${targetRow}.`;
      
      try {
        MailApp.sendEmail({
          to: "owen@keslowcamera.com,ryan@keslowcamera.com,chad@keslowcamera.com",
          subject: emailSubject,
          body: emailBody
        });
        Logger.log('✅ Email notification sent successfully');
      } catch (error) {
        Logger.log(`❌ Failed to send email notification: ${error.toString()}`);
      }
      
      Logger.log('🏁 AMLFFormSubmit completed - new camera added');
      return;
    }
    console.log(`✅ Found matching barcode '${formData[AMLFResponseCOLS.BARCODE]}' in Database.`);
  } else {
    console.log(`✅ Found matching Serial Number '${formData[AMLFResponseCOLS.SERIAL]}' in Database.`);
  }

  const targetRow = dbRowIndex + 2; // Adjust for header
  
  // Insert a snapshot of the found row before making changes
  Logger.log('Snapshot before changes for row ' + targetRow + ': ' + JSON.stringify(dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]));

  // Get current values from database for this row
  var currentValues = dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]; // Get all columns A-P
  var cameraName = currentValues[AMLFDatabaseCOLS.CAMERA - 1];
  var currentSerial = currentValues[AMLFDatabaseCOLS.SERIAL - 1];
  var currentBarcode = currentValues[AMLFDatabaseCOLS.BARCODE - 1];

  // Update database with form data
  if (serviceDate) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.SERVICE).setValue(serviceDate);  // Most Recent Service Date
    console.log(`✅ Updated service date to '${serviceDate}' for row ${targetRow}`);
  }

  if (formData[AMLFResponseCOLS.FIRMWARE]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.FIRMWARE).setValue(formData[AMLFResponseCOLS.FIRMWARE]); // Firmware
    console.log(`✅ Updated firmware to '${formData[AMLFResponseCOLS.FIRMWARE]}' for row ${targetRow}`);
  }

  // Update mount type
  if (formData[AMLFResponseCOLS.MOUNT]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.MOUNT).setValue(formData[AMLFResponseCOLS.MOUNT]);
    console.log(`✅ Updated mount type to '${formData[AMLFResponseCOLS.MOUNT]}' for row ${targetRow}`);
  }

  // Update cage type
  if (formData[AMLFResponseCOLS.CAGE_TYPE]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.CAGE).setValue(formData[AMLFResponseCOLS.CAGE_TYPE]);
    console.log(`✅ Updated cage type to '${formData[AMLFResponseCOLS.CAGE_TYPE]}' for row ${targetRow}`);
  }

  // Update battery plate type
  if (formData[AMLFResponseCOLS.BATTERY_PLATE]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.BATTERY).setValue(formData[AMLFResponseCOLS.BATTERY_PLATE]);
    console.log(`✅ Updated battery plate type to '${formData[AMLFResponseCOLS.BATTERY_PLATE]}' for row ${targetRow}`);
  }

  // Update camera hours
  if (formData[AMLFResponseCOLS.HOURS]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.HOURS).setValue(formData[AMLFResponseCOLS.HOURS]);
    console.log(`✅ Updated hours to '${formData[AMLFResponseCOLS.HOURS]}' for row ${targetRow}`);
  }

  // Update visual impression
  if (formData[AMLFResponseCOLS.VISUAL_RATING]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.VISUAL).setValue(formData[AMLFResponseCOLS.VISUAL_RATING]);
    console.log(`✅ Updated visual to '${formData[AMLFResponseCOLS.VISUAL_RATING]}' for row ${targetRow}`);
  }

  // Update notes
  dbSheet.getRange(targetRow, AMLFDatabaseCOLS.NOTES).setValue(formData[AMLFResponseCOLS.NOTES] || '');  // Notes
  console.log(`✅ Updated notes to '${formData[AMLFResponseCOLS.NOTES] || ''}' for row ${targetRow}`);

  // Update RTR status
  const status = formData[AMLFResponseCOLS.STATUS];
  const normalizedStatus = status ? status.trim().toLowerCase() : "";
  if (normalizedStatus === "ready to rent") {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.STATUS).setValue("RTR");  // RTR Status
    console.log(`✅ Set status to "RTR" for row ${targetRow}`);
  } else if (normalizedStatus === "serviced for order") {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.STATUS).setValue("Pulled");  // Status must be one of: RTR, Shipped, Returned, Pulled, UNKNOWN, Repair
    console.log(`✅ Set status to "Pulled" for row ${targetRow} (converted from "Serviced For Order")`);
  } else if (normalizedStatus === "reserve body") {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.STATUS).setValue("Reserve");  // Status
    console.log(`✅ Set status to "Reserve" for row ${targetRow} (converted from "Reserve Body")`);
  } else if (normalizedStatus.includes("other")) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.STATUS).setValue("UNKNOWN");  // Status
    console.log(`✅ Set status to "UNKNOWN" for row ${targetRow} (converted from "${status}")`);
  } else if (status) {
    // Only set status if it matches one of the allowed values
    const allowedStatuses = ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"];
    if (allowedStatuses.includes(status)) {
      dbSheet.getRange(targetRow, AMLFDatabaseCOLS.STATUS).setValue(status);
      console.log(`✅ Set status to "${status}" for row ${targetRow}`);
      
      // Add repair notification
      if (status === "Repair") {
        cameraRepairRobot(cameraName, currentSerial, currentBarcode, userInfo.fullName);
      }
    } else {
      console.log(`⚠️ Skipped setting invalid status "${status}" - must be one of: ${allowedStatuses.join(", ")}`);
    }
  }

  // Update LAST_SERVICED_BY with email from form response
  if (formData[AMLFResponseCOLS.EMAIL]) {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.LAST_SERVICED_BY).setValue(formData[AMLFResponseCOLS.EMAIL]);
    console.log(`✅ Updated LAST_SERVICED_BY to '${formData[AMLFResponseCOLS.EMAIL]}' for row ${targetRow}`);
  }

  // Update location with user's city
  if (userInfo.city && userInfo.city !== 'City not found') {
    dbSheet.getRange(targetRow, AMLFDatabaseCOLS.LOCATION).setValue(userInfo.city);
    console.log(`✅ Updated location to '${userInfo.city}' for row ${targetRow}`);
  }

  // Insert a snapshot of the found row after making changes
  Logger.log('Snapshot after changes for row ' + targetRow + ': ' + JSON.stringify(dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]));
  
  Logger.log('🏁 AMLFFormSubmit completed - existing camera updated');
} 