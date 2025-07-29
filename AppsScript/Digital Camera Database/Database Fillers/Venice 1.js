function rectifyVenice1Database() {
    // === CONFIGURABLE CONSTANTS ===
    const CAMERA_TYPE = "Sony VENICE 1 HFR Digital Camera";
    const CAMERA_FORM_RESPONSE_SHEET = "Sony Venice 1";
    const CAMERA_DATABASE_SHEET = "Venice 1 Body Status";

    // === DATABASE SHEET COLUMNS ===
    const DB_COLUMNS = {
        QTY: 0,                    // A - Set to 1 for all entries
        CAMERA: 1,                 // B - From BARCODE_DB.EQUIP_NAME
        SERIAL: 2,                 // C - From FORM.SERIAL
        BARCODE: 3,                // D - From FORM.BARCODE
        STATUS: 4,                 // E - IGNORE
        SERVICE_DATE: 5,           // F - From FORM.TIMESTAMP
        LOCATION: 6,               // G - From BARCODE_DB.LOCATION
        OWNER: 7,                  // H - From BARCODE_DB.OWNER
        SENSOR_BLOCK: 8,           // I - From FORM.SENSOR_BLOCK
        MOUNT_TYPE: 9,             // J - IGNORE
        CAMERA_FIRMWARE: 10,       // K - From FORM.CAMERA_FIRMWARE
        R7_FIRMWARE: 11,           // L - From FORM.R7_FIRMWARE
        HOURS: 12,                 // M - From FORM.OPERATING_HOURS
        LAST_SERVICED_BY: 13,      // N - From FORM.EMAIL
        VISUAL: 14,                // O - From FORM.VISUAL_IMPRESSION
        NOTES: 15                  // P - From FORM.TECH_NOTES
    };

    // === FORM RESPONSE SHEET COLUMNS ===
    const FORM_COLUMNS = {
        TIMESTAMP: 0,              // A
        EMAIL: 1,                  // B
        BARCODE: 2,                // C
        SENSOR_BLOCK: 3,           // D
        SERIAL: 4,                 // E
        CLEAN_BODY: 5,             // F
        INSPECT_SCREWS: 6,         // G
        TEST_ROD_LOCKS: 7,         // H
        BASE_PLATE_ALIGN: 8,       // I
        TEST_DOVETAIL: 9,          // J
        INSPECT_DISPLAY: 10,       // K
        VISUAL_IMPRESSION: 11,     // L
        FIRMWARE_START: 12,        // M
        CAMERA_FIRMWARE: 13,       // N
        CHECK_6K_LICENSE: 14,      // O
        CHECK_43_LICENSE: 15,      // P
        TEST_BUTTONS: 16,          // Q
        TEST_POWER: 17,            // R
        TEST_RS_IO: 18,            // S
        TEST_SDI_HDMI: 19,         // T
        TEST_TIMECODE: 20,         // U
        TEST_AUDIO: 21,            // V
        TEST_FILTER: 22,           // W
        CLEAN_EVF: 23,             // X
        TEST_EVF: 24,              // Y
        TEST_DIOPTER: 25,          // Z
        INSPECT_RECORDERS: 26,     // AA
        TEST_PLAYBACK: 27,         // AB
        CLEAN_SENSOR: 28,          // AC
        CLEAN_MOUNT: 29,           // AD
        SET_PRESETS: 30,           // AE
        BLACK_LEVELS: 31,          // AF
        BACK_FOCUS: 32,            // AG
        BACK_FOCUS_CHECK: 33,      // AH
        TECH_NOTES: 34,            // AI
        SET_EVF_DEFAULTS: 35,      // AJ
        CLEAR_LUTS: 36,            // AK
        FACTORY_RESET: 37,         // AL
        R7_FIRMWARE: 38,           // AM
        CHECK_HFR_LICENSE: 39,     // AN
        CHECK_HFR_LICENSE_V4: 40,  // AO
        INSPECT_CONNECTIONS: 41,   // AP
        REMOVE_MEDIA: 42,          // AQ
        OPERATING_HOURS: 43        // AR
    };

    // === BARCODE DATABASE COLUMNS ===
    const BARCODE_DB_COLUMNS = {
        ID: 0,                     // A
        EQUIP_NAME: 1,             // B
        BARCODE: 2,                // C
        CATEGORY: 3,               // D
        STATUS: 4,                 // E
        OWNER: 5,                  // F
        LOCATION: 6                // G
    };

    console.log('Starting database rectification process...');
    const startTime = new Date();
    
    // Helper function to normalize barcodes
    function normalizeBarcode(barcode) {
      if (!barcode) return '';
      return barcode.toString().trim().toUpperCase();
    }

    // Helper function to trim K1 serial numbers
    function trimK1Serial(serial) {
      if (!serial) return '';
      const str = serial.toString().trim();
      if (str.startsWith('K1') && str.includes('-')) {
        return str.split('-')[1];
      }
      return str;
    }
    
    // Get the source sheet with responses
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName(CAMERA_FORM_RESPONSE_SHEET);
    if (!sourceSheet) {
      console.error('Sheet "' + CAMERA_FORM_RESPONSE_SHEET + '" not found!');
      SpreadsheetApp.getUi().alert('Error: Sheet "' + CAMERA_FORM_RESPONSE_SHEET + '" not found!');
      return;
    }
    console.log('Source sheet found:', sourceSheet.getName());
    
    // Get the reference database
    console.log('Opening reference database...');
    const referenceSpreadsheet = SpreadsheetApp.openById("1GBFiSsbKa7MTJzCAsjONvrKCcWVLG0ejLRLWlwOJLag");
    const referenceSheet = referenceSpreadsheet.getSheetByName("Database");
    console.log('Reference database sheet accessed successfully');
    
    // Get all data from reference database
    console.log('Fetching reference database data...');
    const referenceData = referenceSheet.getDataRange().getValues();
    
    // Create a map of valid barcodes (only for CAMERA_TYPE and 'active' or 'Repair' in column E)
    const validBarcodes = new Map();
    referenceData.forEach(row => {
      const itemName = row[1]; // Column B
      const barcode = normalizeBarcode(row[2]); // Column C
      const status = row[4] ? row[4].toString().toLowerCase() : ''; // Column E
      if (itemName === CAMERA_TYPE && barcode && (status.includes('active') || status.includes('Repair'))) {
        validBarcodes.set(barcode, true);
      }
    });
    console.log('Number of valid', CAMERA_TYPE, 'barcodes in database:', validBarcodes.size);
    console.log('First few valid barcodes:', Array.from(validBarcodes.keys()).slice(0, 5));
    
    // Get all data from source sheet
    console.log('Fetching source sheet data...');
    const sourceData = sourceSheet.getDataRange().getValues();
    console.log('Total rows in source sheet:', sourceData.length);
    
    // Create a map to track barcode-serial pair frequencies
    const barcodeSerialFrequencies = new Map();
    let totalBarcodesFound = 0;
    
    // First pass: count frequencies of each barcode-serial pair
    sourceData.forEach((row, index) => {
      if (index === 0) return; // Skip header row
      const originalBarcode = row[FORM_COLUMNS.BARCODE];
      const serial = row[FORM_COLUMNS.SERIAL];
      const normalizedBarcode = normalizeBarcode(originalBarcode);
      
      if (normalizedBarcode !== '') {
        totalBarcodesFound++;
        const key = `${normalizedBarcode}|${serial}`;
        const currentCount = barcodeSerialFrequencies.get(key) || 0;
        barcodeSerialFrequencies.set(key, currentCount + 1);
      }
    });
    
    // Create a map to store the most frequent serial for each barcode
    const mostFrequentPairs = new Map();
    barcodeSerialFrequencies.forEach((frequency, key) => {
      const [barcode, serial] = key.split('|');
      const currentBest = mostFrequentPairs.get(barcode);
      
      if (!currentBest || frequency > currentBest.frequency) {
        mostFrequentPairs.set(barcode, {
          serial: serial,
          frequency: frequency,
          originalBarcode: barcode // Store original barcode for output
        });
      }
    });
    
    console.log('Total barcodes found in source sheet:', totalBarcodesFound);
    console.log('Number of unique barcodes in source sheet:', mostFrequentPairs.size);
    console.log('First few source barcodes:', Array.from(mostFrequentPairs.keys()).slice(0, 5));
    
    // Filter barcodes that exist in reference database and are CAMERA_TYPE
    const validPairs = [];
    let matchCount = 0;
    mostFrequentPairs.forEach((data, normalizedBarcode) => {
      if (validBarcodes.has(normalizedBarcode)) {
        matchCount++;
        validPairs.push([data.serial, data.originalBarcode]); // Use original barcode for output
      } else {
        console.log('No valid match found for barcode:', normalizedBarcode);
      }
    });
    console.log('Number of valid barcode-serial pairs:', validPairs.length);
    console.log('Match count:', matchCount);
    
    // Get or create the target sheet
    let targetSheet = spreadsheet.getSheetByName(CAMERA_DATABASE_SHEET);
    if (!targetSheet) {
      console.log('Creating new target sheet...');
      targetSheet = spreadsheet.insertSheet(CAMERA_DATABASE_SHEET);
      // Add headers only if sheet is new
      targetSheet.getRange(1, 3, 1, 2).setValues([['Serial Number', 'Barcode']]);
      SpreadsheetApp.flush();
    }
    
    // Create a map of barcode database entries
    const barcodeDbMap = new Map();
    referenceData.forEach(row => {
        const barcode = normalizeBarcode(row[BARCODE_DB_COLUMNS.BARCODE]);
        if (barcode) {
            barcodeDbMap.set(barcode, {
                equipName: row[BARCODE_DB_COLUMNS.EQUIP_NAME],
                status: row[BARCODE_DB_COLUMNS.STATUS],
                owner: row[BARCODE_DB_COLUMNS.OWNER],
                location: row[BARCODE_DB_COLUMNS.LOCATION]
            });
        }
    });

    // Create a map of form responses
    const formResponseMap = new Map();
    sourceData.forEach((row, index) => {
        if (index === 0) return; // Skip header row
        const barcode = normalizeBarcode(row[FORM_COLUMNS.BARCODE]);
        if (barcode) {
            formResponseMap.set(barcode, {
                timestamp: row[FORM_COLUMNS.TIMESTAMP],
                serial: row[FORM_COLUMNS.SERIAL],
                email: row[FORM_COLUMNS.EMAIL],
                visualImpression: row[FORM_COLUMNS.VISUAL_IMPRESSION],
                sensorBlock: row[FORM_COLUMNS.SENSOR_BLOCK],
                cameraFirmware: row[FORM_COLUMNS.CAMERA_FIRMWARE],
                r7Firmware: row[FORM_COLUMNS.R7_FIRMWARE],
                operatingHours: row[FORM_COLUMNS.OPERATING_HOURS],
                techNotes: row[FORM_COLUMNS.TECH_NOTES]
            });
        }
    });

    // Prepare data for writing to target sheet
    const rowsToWrite = [];
    validPairs.forEach(([serial, barcode]) => {
        const barcodeData = barcodeDbMap.get(normalizeBarcode(barcode));
        const formData = formResponseMap.get(normalizeBarcode(barcode));
        
        if (barcodeData && formData) {
            const row = new Array(16).fill(''); // Initialize empty row
            
            // Fill in the data
            row[DB_COLUMNS.QTY] = 1;
            row[DB_COLUMNS.CAMERA] = barcodeData.equipName;
            row[DB_COLUMNS.SERIAL] = trimK1Serial(serial);
            row[DB_COLUMNS.BARCODE] = barcode;
            row[DB_COLUMNS.SERVICE_DATE] = formData.timestamp;
            row[DB_COLUMNS.LOCATION] = barcodeData.location;
            row[DB_COLUMNS.OWNER] = barcodeData.owner;
            row[DB_COLUMNS.SENSOR_BLOCK] = formData.sensorBlock;
            row[DB_COLUMNS.CAMERA_FIRMWARE] = formData.cameraFirmware;
            row[DB_COLUMNS.R7_FIRMWARE] = formData.r7Firmware;
            row[DB_COLUMNS.HOURS] = formData.operatingHours;
            row[DB_COLUMNS.LAST_SERVICED_BY] = formData.email;
            row[DB_COLUMNS.VISUAL] = formData.visualImpression;
            row[DB_COLUMNS.NOTES] = formData.techNotes;
            
            rowsToWrite.push(row);
        }
    });

    // Write the data to the target sheet
    if (rowsToWrite.length > 0) {
        console.log('Writing', rowsToWrite.length, 'rows to target sheet...');
        const range = targetSheet.getRange(3, 1, rowsToWrite.length, 16);
        range.setValues(rowsToWrite);
        SpreadsheetApp.flush();
    }

    // Format the sheet
    targetSheet.autoResizeColumns(1, 16);
    SpreadsheetApp.flush();

    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;

    // Show completion message
    const message = `Process complete.\n\n` +
        `Total barcodes found in source: ${totalBarcodesFound}\n` +
        `Unique barcodes processed: ${mostFrequentPairs.size}\n` +
        `Valid ${CAMERA_TYPE} barcodes found: ${validPairs.length}\n` +
        `Rows written to database: ${rowsToWrite.length}\n` +
        `Results written to: ${CAMERA_DATABASE_SHEET} (starting from row 3)\n` +
        `Execution time: ${executionTime.toFixed(2)} seconds`;

    console.log('Process complete. Final statistics:', message);
    SpreadsheetApp.getUi().alert(message);
} 