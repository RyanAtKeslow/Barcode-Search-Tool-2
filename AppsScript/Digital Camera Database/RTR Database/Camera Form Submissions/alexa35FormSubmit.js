const Alexa35DatabaseCOLS = {
  SERIAL: 3,            // Column C
  BARCODE: 4,           // Column D
  STATUS: 5,            // Column E (RTR STATUS)
  SERVICE: 6,           // Column F (Most Recent Service Date)
  LOCATION: 7,          // Column G
  MOUNT: 8,             // Column H (MOUNT TYPE)
  OWNER: 9,             // Column I (OWNER)
  BATTERY: 10,          // Column J (Battery Plate Type)
  FIRMWARE: 11,         // Column K
  HOURS: 12,            // Column L
  NOTES: 13,            // Column M
  LAST_SERVICED_BY: 14, // Column N
  VISUAL: 15            // Column O
};

const Alexa35ResponseCOLS = {
  SERVICE: 0,        // Column A - Most Recent Service
  EMAIL: 1,          // Column B - Email
  BARCODE: 2,        // Column C - Barcode
  SERIAL: 3,         // Column D - Serial Number
  MOUNT: 4,          // Column E - Mount Type
  VISUAL: 8,         // Column I - Visual
  FIRMWARE: 9,       // Column J - Firmware Version
  STATUS: 17,        // Column R - Status
  NOTES: 18          // Column S - Notes
};

function alexa35FormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("NEW Alexa 35");
  const dbSheet = ss.getSheetByName("ALEXA 35 Body Status");

  if (!formSheet || !dbSheet) {
    throw new Error("⚠️ One or both sheets not found. Check sheet names.");
  }

  const formData = e.values;

  // Get user info from email
  const userInfo = FetchUserFromFormSubmitViaEmail(formData[Alexa35ResponseCOLS.EMAIL]);
  console.log(`📧 User info retrieved: ${JSON.stringify(userInfo)}`);

  // Format the service date as M/D/YYYY
  let serviceDate = "";
  if (formData[Alexa35ResponseCOLS.SERVICE]) {
    const date = new Date(formData[Alexa35ResponseCOLS.SERVICE]);
    serviceDate = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    console.log(`Formatted service date from '${formData[Alexa35ResponseCOLS.SERVICE]}' to '${serviceDate}'`);
  }

  // First try to find by serial number in database (Column C)
  let dbData = dbSheet.getRange(2, Alexa35DatabaseCOLS.SERIAL, dbSheet.getLastRow() - 1, 1).getValues(); // Column C
  let dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[Alexa35ResponseCOLS.SERIAL].toString().trim());

  // If serial number not found, try to find by barcode (Column D)
  if (dbRowIndex === -1) {
    console.log(`⚠️ Serial Number '${formData[Alexa35ResponseCOLS.SERIAL]}' not found in Database column C, trying barcode match...`);
    dbData = dbSheet.getRange(2, Alexa35DatabaseCOLS.BARCODE, dbSheet.getLastRow() - 1, 1).getValues(); // Column D
    dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[Alexa35ResponseCOLS.BARCODE].toString().trim());
    
    if (dbRowIndex === -1) {
      console.log(`❌ Neither Serial Number '${formData[Alexa35ResponseCOLS.SERIAL]}' nor Barcode '${formData[Alexa35ResponseCOLS.BARCODE]}' found in Database.`);
      
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
      const newRow = new Array(16).fill(''); // Initialize array with 16 empty strings
      
      newRow[1] = "ARRI ALEXA 35 Camera Body"; // CAMERA
      newRow[Alexa35DatabaseCOLS.SERIAL - 1] = formData[Alexa35ResponseCOLS.SERIAL].toString(); // Ensure serial is treated as text
      newRow[Alexa35DatabaseCOLS.BARCODE - 1] = formData[Alexa35ResponseCOLS.BARCODE].toString(); // Ensure barcode is treated as text
      newRow[Alexa35DatabaseCOLS.SERVICE - 1] = serviceDate;
      newRow[Alexa35DatabaseCOLS.MOUNT - 1] = formData[Alexa35ResponseCOLS.MOUNT];
      newRow[Alexa35DatabaseCOLS.FIRMWARE - 1] = formData[Alexa35ResponseCOLS.FIRMWARE];
      newRow[Alexa35DatabaseCOLS.VISUAL - 1] = formData[Alexa35ResponseCOLS.VISUAL];
      newRow[Alexa35DatabaseCOLS.LAST_SERVICED_BY - 1] = formData[Alexa35ResponseCOLS.EMAIL];
      newRow[Alexa35DatabaseCOLS.LOCATION - 1] = userInfo.city; // Add location from user info
      
      // Set status based on form data
      const status = formData[Alexa35ResponseCOLS.STATUS];
      const normalizedStatus = status ? status.trim().toLowerCase() : "";
      if (normalizedStatus === "ready to rent") {
        newRow[Alexa35DatabaseCOLS.STATUS - 1] = "RTR";
      } else if (normalizedStatus === "serviced for order") {
        newRow[Alexa35DatabaseCOLS.STATUS - 1] = "Pulled";
      } else if (normalizedStatus === "reserve body") {
        newRow[Alexa35DatabaseCOLS.STATUS - 1] = "Reserve";
      } else if (normalizedStatus.includes("other")) {
        newRow[Alexa35DatabaseCOLS.STATUS - 1] = "UNKNOWN";
      } else if (status && ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"].includes(status)) {
        newRow[Alexa35DatabaseCOLS.STATUS - 1] = status;
      }
      
      // Write the new row
      dbSheet.getRange(targetRow, 1, 1, 16).setValues([newRow]);
      
      // Format serial and barcode columns as text to preserve leading zeros
      dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.SERIAL).setNumberFormat('@');
      dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.BARCODE).setNumberFormat('@');
      
      console.log(`✅ Added new camera entry at row ${targetRow} with serial: ${formData[Alexa35ResponseCOLS.SERIAL]}, barcode: ${formData[Alexa35ResponseCOLS.BARCODE]}`);
      
      // Send email notification
      const emailSubject = "Camera Not Found in Database - Alexa 35";
      const emailBody = `A camera was submitted via service form and could not find a match in the database.\n\n` +
                       `Camera Type: Alexa 35\n` +
                       `Serial Number: ${formData[Alexa35ResponseCOLS.SERIAL]}\n` +
                       `Barcode: ${formData[Alexa35ResponseCOLS.BARCODE]}\n\n` +
                       `Please confirm the accuracy of this information.\n\n` +
                       `The camera has been added to the database at row ${targetRow}.`;
      
      MailApp.sendEmail({
        to: "owen@keslowcamera.com, ryan@keslowcamera.com, chad@keslowcamera.com",
        subject: emailSubject,
        body: emailBody
      });
      
      return;
    }
    console.log(`✅ Found matching barcode '${formData[Alexa35ResponseCOLS.BARCODE]}' in Database.`);
  } else {
    console.log(`✅ Found matching Serial Number '${formData[Alexa35ResponseCOLS.SERIAL]}' in Database.`);
  }

  const targetRow = dbRowIndex + 2; // Adjust for header

  // Get all current values from database for this row
  const currentValues = dbSheet.getRange(targetRow, 1, 1, 15).getValues()[0]; // Get all columns A-O
  const cameraName = currentValues[1]; // Column B - Camera name
  const currentSerial = currentValues[Alexa35DatabaseCOLS.SERIAL - 1];
  const currentBarcode = currentValues[Alexa35DatabaseCOLS.BARCODE - 1];
  const oldMount = currentValues[Alexa35DatabaseCOLS.MOUNT - 1];
  const newMount = formData[Alexa35ResponseCOLS.MOUNT];

  // Update database with form data
  if (serviceDate) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.SERVICE).setValue(serviceDate);  // Most Recent Service
    console.log(`✅ Updated service date to '${serviceDate}' for row ${targetRow}`);
  }

  if (formData[Alexa35ResponseCOLS.FIRMWARE]) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.FIRMWARE).setValue(formData[Alexa35ResponseCOLS.FIRMWARE]); // Firmware
    console.log(`✅ Updated firmware to '${formData[Alexa35ResponseCOLS.FIRMWARE]}' for row ${targetRow}`);
  }

  // Update notes
  dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.NOTES).setValue(formData[Alexa35ResponseCOLS.NOTES] || '');  // Notes
  console.log(`✅ Updated notes to '${formData[Alexa35ResponseCOLS.NOTES] || ''}' for row ${targetRow}`);

  // Update RTR status
  const status = formData[Alexa35ResponseCOLS.STATUS];
  const normalizedStatus = status ? status.trim().toLowerCase() : "";
  if (normalizedStatus === "ready to rent") {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.STATUS).setValue("RTR");  // RTR Status
    console.log(`✅ Set status to "RTR" for row ${targetRow}`);
  } else if (normalizedStatus === "serviced for order") {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.STATUS).setValue("Pulled");  // Status must be one of: RTR, Shipped, Returned, Pulled, UNKNOWN, Repair
    console.log(`✅ Set status to "Pulled" for row ${targetRow} (converted from "Serviced For Order")`);
  } else if (normalizedStatus === "reserve body") {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.STATUS).setValue("Reserve");  // Status
    console.log(`✅ Set status to "Reserve" for row ${targetRow} (converted from "Reserve Body")`);
  } else if (normalizedStatus.includes("other")) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.STATUS).setValue("UNKNOWN");  // Status
    console.log(`✅ Set status to "UNKNOWN" for row ${targetRow} (converted from "${status}")`);
  } else if (status) {
    // Only set status if it matches one of the allowed values
    const allowedStatuses = ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"];
    if (allowedStatuses.includes(status)) {
      dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.STATUS).setValue(status);
      console.log(`✅ Set status to "${status}" for row ${targetRow}`);
      
      // Add repair notification
      if (status === "Repair") {
        const cameraName = "ALEXA 35"; // Alexa 35 is the camera name
        const cameraSN = dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.SERIAL).getValue();
        const cameraBC = dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.BARCODE).getValue();
        cameraRepairRobot(cameraName, cameraSN, cameraBC, userInfo.fullName);
      }
    } else {
      console.log(`⚠️ Skipped setting invalid status "${status}" - must be one of: ${allowedStatuses.join(", ")}`);
    }
  }

  // Update LAST_SERVICED_BY with email from form response
  if (formData[Alexa35ResponseCOLS.EMAIL]) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.LAST_SERVICED_BY).setValue(formData[Alexa35ResponseCOLS.EMAIL]);
    console.log(`✅ Updated LAST_SERVICED_BY to '${formData[Alexa35ResponseCOLS.EMAIL]}' for row ${targetRow}`);
  }

  // Update VISUAL (column I in response, column O in database)
  if (formData[Alexa35ResponseCOLS.VISUAL]) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.VISUAL).setValue(formData[Alexa35ResponseCOLS.VISUAL]);
    console.log(`✅ Updated VISUAL to '${formData[Alexa35ResponseCOLS.VISUAL]}' for row ${targetRow}`);
  }

  // Mount change detection and notification
  if (oldMount && newMount && oldMount !== newMount) {
    lensMountRobot(cameraName, currentSerial, currentBarcode, oldMount, newMount, userInfo.fullName);
    Logger.log(`Sent lens mount change notification: ${cameraName} SN(${currentSerial}) BC(${currentBarcode}) changed from ${oldMount} -> ${newMount} by ${userInfo.fullName}`);
  }
  
  // Update mount type
  if (formData[Alexa35ResponseCOLS.MOUNT]) {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.MOUNT).setValue(formData[Alexa35ResponseCOLS.MOUNT]);
    console.log(`✅ Updated mount type to '${formData[Alexa35ResponseCOLS.MOUNT]}' for row ${targetRow}`);
  }

  // Update location with user's city
  if (userInfo.city && userInfo.city !== 'City not found') {
    dbSheet.getRange(targetRow, Alexa35DatabaseCOLS.LOCATION).setValue(userInfo.city);
    console.log(`✅ Updated location to '${userInfo.city}' for row ${targetRow}`);
  }
} 