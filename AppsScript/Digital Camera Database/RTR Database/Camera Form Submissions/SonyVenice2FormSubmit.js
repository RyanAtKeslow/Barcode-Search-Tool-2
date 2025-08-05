const Venice2DatabaseCOLS = {
  CAMERA: 2,                // Column B
  SERIAL: 3,                // Column C
  BARCODE: 4,               // Column D
  STATUS: 5,                // Column E
  SERVICE: 6,               // Column F (Most Recent Service Date)
  LOCATION: 7,              // Column G
  OWNER: 8,                 // Column H (new)
  SENSOR_BLOCK: 9,          // Column I
  SENSOR_BLOCK_BARCODE: 10, // Column J
  MOUNT: 11,                // Column K
  FIRMWARE: 12,             // Column L
  HOURS: 13,                // Column M
  LAST_SERVICED_BY: 14,     // Column N
  VISUAL: 15,               // Column O
  NOTES: 16                 // Column P
};

const Venice2ResponseCOLS = {
  TIMESTAMP: 0,         // Column A
  EMAIL: 1,             // Column B
  BARCODE: 2,           // Column C
  SERIAL: 3,            // Column D
  SENSOR_BLOCK: 4,      // Column E
  SENSOR_RES: 5,        // Column F
  LENS_MOUNT: 6,        // Column G
  SERVICE_TYPE: 7,      // Column H
  VISUAL: 11,           // Column L
  FIRMWARE: 12,         // Column M
  HOURS: 13,            // Column N
  LICENSES: 14,         // Column O
  STATUS: 20,           // Column U
  NOTES: 21             // Column V
};

function SonyVenice2FormSubmit(e) {
  Logger.log('🎬 Starting SonyVenice2FormSubmit');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("NEW Sony Venice 2");
  const dbSheet = ss.getSheetByName("VENICE 2 Body Status");

  if (!formSheet || !dbSheet) {
    Logger.log('❌ Sheet validation failed');
    throw new Error("⚠️ One or both sheets not found. Check sheet names.");
  }

  const formData = e.values;
  Logger.log(`📊 Form data received: Serial=${formData[Venice2ResponseCOLS.SERIAL]}, Barcode=${formData[Venice2ResponseCOLS.BARCODE]}, Email=${formData[Venice2ResponseCOLS.EMAIL]}`);

  // Debug: Print all serial numbers and barcodes from the database
  const allSerials = dbSheet.getRange(2, Venice2DatabaseCOLS.SERIAL, dbSheet.getLastRow() - 1, 1).getValues().map(row => row[0]);
  const allBarcodes = dbSheet.getRange(2, Venice2DatabaseCOLS.BARCODE, dbSheet.getLastRow() - 1, 1).getValues().map(row => row[0]);
  Logger.log('All serial numbers in DB: ' + JSON.stringify(allSerials));
  Logger.log('All barcodes in DB: ' + JSON.stringify(allBarcodes));

  // Format the service date as M/D/YYYY (if present in the form, e.g., TIMESTAMP)
  let serviceDate = "";
  if (formData[Venice2ResponseCOLS.TIMESTAMP]) {
    const date = new Date(formData[Venice2ResponseCOLS.TIMESTAMP]);
    serviceDate = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    console.log(`Formatted service date from '${formData[Venice2ResponseCOLS.TIMESTAMP]}' to '${serviceDate}'`);
  }

  // First try to find by serial number in database (Column C)
  let dbData = dbSheet.getRange(2, Venice2DatabaseCOLS.SERIAL, dbSheet.getLastRow() - 1, 1).getValues(); // Column C
  let dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[Venice2ResponseCOLS.SERIAL].toString().trim());

  Logger.log('dbData (serials): ' + JSON.stringify(dbData));
  Logger.log('Form serial: [' + formData[Venice2ResponseCOLS.SERIAL] + ']');
  Logger.log('dbRowIndex (serial): ' + dbRowIndex);

  // If serial number not found, try to find by barcode (Column D)
  if (dbRowIndex === -1) {
    console.log(`⚠️ Serial Number '${formData[Venice2ResponseCOLS.SERIAL]}' not found in Database column C, trying barcode match...`);
    dbData = dbSheet.getRange(2, Venice2DatabaseCOLS.BARCODE, dbSheet.getLastRow() - 1, 1).getValues(); // Column D
    dbRowIndex = dbData.findIndex(row => row[0].toString().trim() === formData[Venice2ResponseCOLS.BARCODE].toString().trim());
    
    if (dbRowIndex === -1) {
      console.log(`❌ Neither Serial Number '${formData[Venice2ResponseCOLS.SERIAL]}' nor Barcode '${formData[Venice2ResponseCOLS.BARCODE]}' found in Database.`);
      
      // Find first empty row instead of appending to the end
      let targetRow = 3; // Start from row 3 (after header and potential first data row)
      let maxRows = dbSheet.getMaxRows();
      
      // Find first empty row by checking if CAMERA column is empty
      for (let i = 3; i <= maxRows; i++) {
        let cellValue = dbSheet.getRange(i, Venice2DatabaseCOLS.CAMERA).getValue();
        if (!cellValue || cellValue.toString().trim() === '') {
          targetRow = i;
          break;
        }
      }
      
      // If no empty row found, add to the end
      if (targetRow === 3 && dbSheet.getRange(3, Venice2DatabaseCOLS.CAMERA).getValue()) {
        targetRow = dbSheet.getLastRow() + 1;
      }
      
      // Get user info for the new entry
      const userInfo = FetchUserFromFormSubmitViaEmail(formData[Venice2ResponseCOLS.EMAIL]);
      Logger.log(`User info for new entry: ${JSON.stringify(userInfo)}`);
      
      // Prepare the new row data
      const newRow = new Array(17).fill(''); // Initialize array with 17 empty strings
      
      newRow[1] = "Sony VENICE 2 Digital Camera"; // CAMERA
      newRow[Venice2DatabaseCOLS.SERIAL - 1] = formData[Venice2ResponseCOLS.SERIAL].toString(); // Ensure serial is treated as text
      newRow[Venice2DatabaseCOLS.BARCODE - 1] = formData[Venice2ResponseCOLS.BARCODE].toString(); // Ensure barcode is treated as text
      newRow[Venice2DatabaseCOLS.SERVICE - 1] = serviceDate;
      newRow[Venice2DatabaseCOLS.LOCATION - 1] = userInfo.city;
      newRow[Venice2DatabaseCOLS.OWNER - 1] = "";
      newRow[Venice2DatabaseCOLS.SENSOR_BLOCK - 1] = formData[Venice2ResponseCOLS.SENSOR_RES];
      newRow[Venice2DatabaseCOLS.SENSOR_BLOCK_BARCODE - 1] = formData[Venice2ResponseCOLS.SENSOR_BLOCK];
      newRow[Venice2DatabaseCOLS.MOUNT - 1] = formData[Venice2ResponseCOLS.LENS_MOUNT];
      newRow[Venice2DatabaseCOLS.FIRMWARE - 1] = formData[Venice2ResponseCOLS.FIRMWARE]?.replace(/^V/i, '');
      newRow[Venice2DatabaseCOLS.HOURS - 1] = formData[Venice2ResponseCOLS.HOURS];
      newRow[Venice2DatabaseCOLS.VISUAL - 1] = formData[Venice2ResponseCOLS.VISUAL];
      newRow[Venice2DatabaseCOLS.LAST_SERVICED_BY - 1] = formData[Venice2ResponseCOLS.EMAIL];
      
      // Set status based on form data
      const status = formData[Venice2ResponseCOLS.STATUS];
      const normalizedStatus = status ? status.trim().toLowerCase() : "";
      if (normalizedStatus === "ready to rent") {
        newRow[Venice2DatabaseCOLS.STATUS - 1] = "RTR";
      } else if (normalizedStatus === "serviced for order") {
        newRow[Venice2DatabaseCOLS.STATUS - 1] = "Pulled";
      } else if (normalizedStatus === "reserve body") {
        newRow[Venice2DatabaseCOLS.STATUS - 1] = "Reserve";
      } else if (normalizedStatus.includes("other")) {
        newRow[Venice2DatabaseCOLS.STATUS - 1] = "UNKNOWN";
      } else if (status && ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"].includes(status)) {
        newRow[Venice2DatabaseCOLS.STATUS - 1] = status;
      }
      
      // Write the new row
      dbSheet.getRange(targetRow, 1, 1, 17).setValues([newRow]);
      
      // Format serial and barcode columns as text to preserve leading zeros
      dbSheet.getRange(targetRow, Venice2DatabaseCOLS.SERIAL).setNumberFormat('@');
      dbSheet.getRange(targetRow, Venice2DatabaseCOLS.BARCODE).setNumberFormat('@');
      
      console.log(`✅ Added new camera entry at row ${targetRow} with serial: ${formData[Venice2ResponseCOLS.SERIAL]}, barcode: ${formData[Venice2ResponseCOLS.BARCODE]}`);
      
      // Send email notification
      Logger.log('📧 Sending email notification for new camera entry');
      const emailSubject = "Camera Not Found in Database - Sony Venice 2";
      const emailBody = `A camera was submitted via service form and could not find a match in the database.\n\n` +
                       `Camera Type: Sony Venice 2\n` +
                       `Serial Number: ${formData[Venice2ResponseCOLS.SERIAL]}\n` +
                       `Barcode: ${formData[Venice2ResponseCOLS.BARCODE]}\n` +
                       `Submitted by: ${formData[Venice2ResponseCOLS.EMAIL]}\n\n` +
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
      
      Logger.log('🏁 SonyVenice2FormSubmit completed - new camera added');
      return;
    }
    console.log(`✅ Found matching barcode '${formData[Venice2ResponseCOLS.BARCODE]}' in Database.`);
  } else {
    console.log(`✅ Found matching Serial Number '${formData[Venice2ResponseCOLS.SERIAL]}' in Database.`);
  }

  
  const targetRow = dbRowIndex + 2; // Adjust for header
  // Insert a snapshot of the found row before making changes
  Logger.log('Snapshot before changes for row ' + targetRow + ': ' + JSON.stringify(dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]));

  // Get all current values from database for this row
  var currentValues = dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]; // Get all columns A-P
  var cameraName = currentValues[Venice2DatabaseCOLS.CAMERA - 1];
  var currentSerial = currentValues[Venice2DatabaseCOLS.SERIAL - 1];
  var currentBarcode = currentValues[Venice2DatabaseCOLS.BARCODE - 1];
  var oldMount = currentValues[Venice2DatabaseCOLS.MOUNT - 1];
  var oldSensorBlock = currentValues[Venice2DatabaseCOLS.SENSOR_BLOCK - 1];
  var oldSensorBlockBarcode = currentValues[Venice2DatabaseCOLS.SENSOR_BLOCK_BARCODE - 1];
  var email = formData[Venice2ResponseCOLS.EMAIL];
  var userInfo = FetchUserFromFormSubmitViaEmail(email) || { fullName: 'user not found' };
  var newMount = formData[Venice2ResponseCOLS.LENS_MOUNT];
  
  // Check for sensor block barcode match with different sensor modes
  if (formData[Venice2ResponseCOLS.SENSOR_BLOCK] && 
      formData[Venice2ResponseCOLS.SENSOR_BLOCK] === oldSensorBlockBarcode && 
      formData[Venice2ResponseCOLS.SENSOR_RES] !== oldSensorBlock) {
    const sensorModeMsg = `Double checking that ${cameraName} SN(${currentSerial}) BC(${currentBarcode}) should be listed as ${formData[Venice2ResponseCOLS.SENSOR_RES]} mode - ${userInfo.fullName}`;
    venice2SensorRobot(sensorModeMsg);
    Logger.log('Sensor mode verification message sent to GChat: ' + sensorModeMsg);
  }
  
  // Define allowed mount types
  const allowedMountTypes = ['PL', 'LPL Only', 'LPL with PL Adaptor'];
  
  // Handle non-standard mount types
  if (newMount && !allowedMountTypes.includes(newMount)) {
    // Get existing notes
    let existingNotes = dbSheet.getRange(targetRow, Venice2DatabaseCOLS.NOTES).getValue() || '';
    // Add mount information to notes
    let mountNote = `Mount = ${newMount}`;
    let updatedNotes = existingNotes ? `${existingNotes}\n${mountNote}` : mountNote;
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.NOTES).setValue(updatedNotes);
    
    // Set mount to "Other"
    newMount = "Other";
    
    // Send chat notification about non-standard mount using database values
    const mountDisplayValue = `Other (Original input: ${formData[Venice2ResponseCOLS.LENS_MOUNT]})`;
    lensMountRobot(cameraName, currentSerial, currentBarcode, oldMount, mountDisplayValue, userInfo.fullName);
    Logger.log(`Sent lens mount change notification: ${cameraName} SN(${currentSerial}) BC(${currentBarcode}) changed from ${oldMount} -> ${mountDisplayValue} by ${userInfo.fullName}`);
  }
  
  // Update location with user's city
  if (userInfo && userInfo.city) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.LOCATION).setValue(userInfo.city);
    console.log(`✅ Updated location to '${userInfo.city}' for row ${targetRow}`);
  }
  
  var trigger = false;
  if (oldMount && newMount && oldMount !== newMount) {
    // Only trigger for PL <-> LPL Only or PL <-> LPL with PL Adaptor
    var pl = 'PL';
    var lplOnly = 'LPL Only';
    var lplWithPL = 'LPL with PL Adaptor';
    if (
      (oldMount === pl && (newMount === lplOnly || newMount === lplWithPL)) ||
      ((oldMount === lplOnly || oldMount === lplWithPL) && newMount === pl)
    ) {
      trigger = true;
    }
  }
  if (trigger) {
    lensMountRobot(cameraName, currentSerial, currentBarcode, oldMount, newMount, userInfo.fullName);
    Logger.log(`Sent lens mount change notification: ${cameraName} SN(${currentSerial}) BC(${currentBarcode}) changed from ${oldMount} -> ${newMount} by ${userInfo.fullName}`);
  }
  if (newMount) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.MOUNT).setValue(newMount);    // Mount Type
    Logger.log(`✅ Updated mount type to '${newMount}' for row ${targetRow}`);
  }

  // --- Sensor block type change notification logic ---
  var newSensorBlock = formData[Venice2ResponseCOLS.SENSOR_RES];
  if ((oldSensorBlock === '6K' && newSensorBlock === '8K') || (oldSensorBlock === '8K' && newSensorBlock === '6K')) {
    let sensorMsg = '';
    
    if (oldSensorBlock === '8K' && newSensorBlock === '6K') {
      // When switching from 8K to 6K, find the Venice 1 information
      const venice1Sheet = ss.getSheetByName("Venice 1 Body Status");
      if (venice1Sheet) {
        const venice1Data = venice1Sheet.getRange(2, 9, venice1Sheet.getLastRow() - 1, 1).getValues(); // Column I
        const venice1RowIndex = venice1Data.findIndex(row => row[0].toString().trim() === formData[Venice2ResponseCOLS.SENSOR_BLOCK].toString().trim());
        
        if (venice1RowIndex !== -1) {
          const venice1Row = venice1RowIndex + 2; // Adjust for header
          const venice1SerialNum = venice1Sheet.getRange(venice1Row, 3).getValue(); // Column C
          const venice1Barcode = venice1Sheet.getRange(venice1Row, 4).getValue(); // Column D
          
          // Set Venice 1 status to "Inactive"
          venice1Sheet.getRange(venice1Row, 5).setValue("Inactive"); // Column E
          
          // Add note to Venice 1
          const currentDate = new Date();
          const formattedDate = `${currentDate.getMonth() + 1}/${currentDate.getDate()}/${currentDate.getFullYear()}`;
          const noteMessage = `6k Sensor paired with Sony Venice 2 SN ${currentSerial} BC ${currentBarcode} on ${formattedDate} by ${userInfo.fullName}`;
          venice1Sheet.getRange(venice1Row, 16).setValue(noteMessage); // Column P
          
          sensorMsg = `${cameraName} SN(${currentSerial}) BC(${currentBarcode}) has been switched to 6K mode. Sensor block (${formData[Venice2ResponseCOLS.SENSOR_BLOCK]}) was removed from Sony Venice 1 SN(${venice1SerialNum}) BC(${venice1Barcode}) by ${userInfo.fullName}. Venice 1 SN ${venice1SerialNum} BC ${venice1Barcode} is now inactive.`;
        } else {
          sensorMsg = `${cameraName} SN(${currentSerial}) BC(${currentBarcode}) has been switched to 6K mode. Sensor block (${formData[Venice2ResponseCOLS.SENSOR_BLOCK]}) was not found in Venice 1 database. Please check with ${userInfo.fullName} as to which Venice 1 should be downed, if any. `;
        }
      } else {
        sensorMsg = `${cameraName} SN(${currentSerial}) BC(${currentBarcode}) has been switched to 6K mode. Venice 1 database not found. Please check with ${userInfo.fullName} as to which Venice 1 should be downed, if any. `;
      }
    } else {
      // When switching from 6K to 8K
      sensorMsg = `${cameraName} SN(${currentSerial}) BC(${currentBarcode}) has been switched to 8K mode and has sensor block (${formData[Venice2ResponseCOLS.SENSOR_BLOCK]}) attached to it by ${userInfo.fullName}. 6K sensor block BC(${oldSensorBlockBarcode}) should be freed up now`;
    }
    
    venice2SensorRobot(sensorMsg);
    Logger.log('Sensor block change notification sent to GChat.');
  }

  // Update database with form data
  if (serviceDate) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.SERVICE).setValue(serviceDate);  // Most Recent Service Date
    console.log(`✅ Updated service date to '${serviceDate}' for row ${targetRow}`);
  }

  if (formData[Venice2ResponseCOLS.FIRMWARE]) {
    // Normalize firmware version by removing 'V' prefix
    const normalizedFirmware = formData[Venice2ResponseCOLS.FIRMWARE].replace(/^V/i, '');
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.FIRMWARE).setValue(normalizedFirmware); // Firmware
    console.log(`✅ Updated firmware to '${normalizedFirmware}' for row ${targetRow}`);
  }

  // Write Sensor Block Barcode from form (column E) to database (column H)
  if (formData[Venice2ResponseCOLS.SENSOR_BLOCK]) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.SENSOR_BLOCK_BARCODE).setValue(formData[Venice2ResponseCOLS.SENSOR_BLOCK]);
    console.log(`✅ Updated sensor block barcode to '${formData[Venice2ResponseCOLS.SENSOR_BLOCK]}' for row ${targetRow}`);
  }

  // Update Sensor Resolution from form (column F) to database (column I)
  if (formData[Venice2ResponseCOLS.SENSOR_RES]) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.SENSOR_BLOCK).setValue(formData[Venice2ResponseCOLS.SENSOR_RES]);
    console.log(`✅ Updated sensor resolution to '${formData[Venice2ResponseCOLS.SENSOR_RES]}' for row ${targetRow}`);
  }

  // Update HOURS
  if (formData[Venice2ResponseCOLS.HOURS]) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.HOURS).setValue(formData[Venice2ResponseCOLS.HOURS]);
    console.log(`✅ Updated hours to '${formData[Venice2ResponseCOLS.HOURS]}' for row ${targetRow}`);
  }

  // Update VISUAL
  if (formData[Venice2ResponseCOLS.VISUAL]) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.VISUAL).setValue(formData[Venice2ResponseCOLS.VISUAL]);
    console.log(`✅ Updated visual to '${formData[Venice2ResponseCOLS.VISUAL]}' for row ${targetRow}`);
  }

  // Update notes (now column N)
  dbSheet.getRange(targetRow, Venice2DatabaseCOLS.NOTES).setValue(formData[Venice2ResponseCOLS.NOTES] || '');
  console.log(`✅ Updated notes to '${formData[Venice2ResponseCOLS.NOTES] || ''}' for row ${targetRow}`);

  // Update RTR status
  const status = formData[Venice2ResponseCOLS.STATUS];
  const normalizedStatus = status ? status.trim().toLowerCase() : "";
  if (normalizedStatus === "ready to rent") {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.STATUS).setValue("RTR");  // RTR Status
    console.log(`✅ Set status to "RTR" for row ${targetRow}`);
  } else if (normalizedStatus === "serviced for order") {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.STATUS).setValue("Pulled");  // Status must be one of: RTR, Shipped, Returned, Pulled, UNKNOWN, Repair
    console.log(`✅ Set status to "Pulled" for row ${targetRow} (converted from "Serviced For Order")`);
  } else if (normalizedStatus === "reserve body") {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.STATUS).setValue("Reserve");  // Status
    console.log(`✅ Set status to "Reserve" for row ${targetRow} (converted from "Reserve Body")`);
  } else if (normalizedStatus.includes("other")) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.STATUS).setValue("UNKNOWN");  // Status
    console.log(`✅ Set status to "UNKNOWN" for row ${targetRow} (converted from "${status}")`);
  } else if (status) {
    // Only set status if it matches one of the allowed values
    const allowedStatuses = ["RTR", "Shipped", "Returned", "Pulled", "UNKNOWN", "Repair", "Reserve"];
    if (allowedStatuses.includes(status)) {
      dbSheet.getRange(targetRow, Venice2DatabaseCOLS.STATUS).setValue(status);
      console.log(`✅ Set status to "${status}" for row ${targetRow}`);
      
      // Add repair notification
      if (status === "Repair") {
        const cameraName = currentValues[Venice2DatabaseCOLS.CAMERA - 1];
        const cameraSN = currentValues[Venice2DatabaseCOLS.SERIAL - 1];
        const cameraBC = currentValues[Venice2DatabaseCOLS.BARCODE - 1];
        cameraRepairRobot(cameraName, cameraSN, cameraBC, userInfo.fullName);
      }
    } else {
      console.log(`⚠️ Skipped setting invalid status "${status}" - must be one of: ${allowedStatuses.join(", ")}`);
    }
  }

  // Update LAST_SERVICED_BY with email from form response
  if (formData[Venice2ResponseCOLS.EMAIL]) {
    dbSheet.getRange(targetRow, Venice2DatabaseCOLS.LAST_SERVICED_BY).setValue(formData[Venice2ResponseCOLS.EMAIL]);
    Logger.log(`✅ Updated LAST_SERVICED_BY to '${formData[Venice2ResponseCOLS.EMAIL]}' for row ${targetRow}`);
  }

  // Insert a snapshot of the found row after making changes
  Logger.log('Snapshot after changes for row ' + targetRow + ': ' + JSON.stringify(dbSheet.getRange(targetRow, 1, 1, 16).getValues()[0]));
  
  Logger.log('🏁 SonyVenice2FormSubmit completed - existing camera updated');
}

 