function rectifyMiniLfDatabase() {
  // TESTING GITHUB FOR RYAN //
    // === CONFIGURABLE CONSTANTS ===
    const CAMERA_TYPE = "ARRI ALEXA Mini LF Camera Body";
    const CAMERA_FORM_RESPONSE_SHEET = "ARRI ALEXA MINI LF";
    const CAMERA_DATABASE_SHEET = "Alexa Mini LF Body Status";

    // === VALID CAGE TYPES ===
    const VALID_CAGE_TYPES = ['Arri', 'Keslow', 'Other', 'Tilta'];

    // === DATABASE SHEET COLUMNS ===
    const DB_COLUMNS = {
        QTY: 0,                    // A - Set to 1 for all entries
        CAMERA: 1,                 // B - From BARCODE_DB.EQUIP_NAME
        SERIAL: 2,                 // C - From FORM.SERIAL
        BARCODE: 3,                // D - From FORM.BARCODE
        RTR_STATUS: 4,             // E - IGNORE
        SERVICE_DATE: 5,           // F - From FORM.TIMESTAMP
        LOCATION: 6,               // G - From BARCODE_DB.LOCATION
        MOUNT_TYPE: 7,             // H - IGNORE
        OWNER: 8,                  // I - From BARCODE_DB.OWNER
        CAGE_TYPE: 9,              // J - From FORM.CAGE_TYPE
        BATTERY_PLATE: 10,         // K - No direct mapping
        FIRMWARE: 11,              // L - From FORM.FIRMWARE_VERSION
        HOURS: 12,                 // M - From FORM.OPERATING_HOURS
        NOTES: 13,                 // N - From FORM.TECH_NOTES
        LAST_SERVICED_BY: 14,      // O - From FORM.EMAIL
        VISUAL: 15                 // P - From FORM.VISUAL_IMPRESSION
    };

    // === FORM RESPONSE SHEET COLUMNS ===
    const FORM_COLUMNS = {
        TIMESTAMP: 0,              // A
        EMAIL: 1,                  // B
        BARCODE: 2,                // C
        SERIAL: 3,                 // D
        CLEAN_BODY: 4,             // E
        INSPECT_SCREWS: 5,         // F
        VISUAL_IMPRESSION: 6,      // G
        CAGE_INSTALLED: 7,         // H
        BASE_PLATE_ALIGN_1: 8,     // I
        TEST_ROD_LOCKS_1: 9,       // J
        TEST_DOVETAIL_1: 10,       // K
        BASE_PLATE_ALIGN_2: 11,    // L
        TEST_ROD_LOCKS_2: 12,      // M
        TEST_DOVETAIL_2: 13,       // N
        TEST_FAN: 14,              // O
        TEST_ONBOARD_POWER: 15,    // P
        TEST_6PIN_POWER: 16,       // Q
        TEST_RS_POWER: 17,         // R
        TEST_DIST_AMP: 18,         // S
        CAGE_TYPE: 19,             // T
        BASE_PLATE_ALIGN_3: 20,    // U
        TEST_ROD_LOCKS_3: 21,      // V
        TEST_DOVETAIL_3: 22,       // W
        TEST_POWER_PORTS: 23,      // X
        BASE_PLATE_ALIGN_4: 24,    // Y
        TEST_ROD_LOCKS_4: 25,      // Z
        TEST_DOVETAIL_4: 26,       // AA
        FIRMWARE_VERSION: 27,      // AB
        TEST_WIFI: 28,             // AC
        TEST_BUTTONS: 29,          // AD
        TEST_ND_FILTER: 30,        // AE
        TEST_8PIN_POWER: 31,       // AF
        TEST_ONBOARD_POWER_2: 32,  // AG
        TEST_RS_IO: 33,            // AH
        TEST_SDI: 34,              // AI
        TEST_AUDIO: 35,            // AJ
        TEST_TIMECODE: 36,         // AK
        TEST_INTERNAL_BATTERY: 37, // AL
        CLEAN_EVF: 38,             // AM
        TEST_DIOPTER: 39,          // AN
        TEST_EVF_BRACKET: 40,      // AO
        TEST_EVF_CABLES: 41,       // AP
        TEST_EVF_BUTTONS: 42,      // AQ
        TEST_MEDIA_DOOR: 43,       // AR
        TEST_CODEX: 44,            // AS
        TEST_PLAYBACK: 45,         // AT
        CLEAN_SENSOR: 46,          // AU
        CLEAN_LPL_MOUNT: 47,       // AV
        CLEAN_PL_ADAPTER: 48,      // AW
        PL_MOUNT_BARCODE: 49,      // AX
        FACTORY_RESET: 50,         // AY
        SET_PRESETS: 51,           // AZ
        BACK_FOCUS: 52,            // BA
        BACK_FOCUS_CHECK: 53,      // BB
        TECH_NOTES: 54,            // BC
        REMOVE_MEDIA: 55,          // BD
        COLOR_TEMP_DRIFT: 56,      // BE
        OPERATING_HOURS: 57        // BF
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
    
    // Helper function to normalize cage type
    function normalizeCageType(cageType) {
        if (!cageType) return '';
        
        const trimmed = cageType.toString().trim();
        console.log('Raw cage type from form:', trimmed);
        
        // Just return the value if it's in our valid types array
        const result = VALID_CAGE_TYPES.includes(trimmed) ? trimmed : '';
        console.log('Normalized result:', result);
        return result;
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
      const originalBarcode = row[2]; // Column C
      const serial = row[3]; // Column D
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
            console.log('Form row', index, 'cage type:', row[FORM_COLUMNS.CAGE_INSTALLED]);
            formResponseMap.set(barcode, {
                timestamp: row[FORM_COLUMNS.TIMESTAMP],
                serial: row[FORM_COLUMNS.SERIAL],
                email: row[FORM_COLUMNS.EMAIL],
                visualImpression: row[FORM_COLUMNS.VISUAL_IMPRESSION],
                cageType: row[FORM_COLUMNS.CAGE_INSTALLED],
                firmwareVersion: row[FORM_COLUMNS.FIRMWARE_VERSION],
                operatingHours: row[FORM_COLUMNS.OPERATING_HOURS],
                techNotes: row[FORM_COLUMNS.TECH_NOTES],
                lplMount: row[FORM_COLUMNS.CLEAN_LPL_MOUNT]
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
            row[DB_COLUMNS.CAGE_TYPE] = normalizeCageType(formData.cageType);
            row[DB_COLUMNS.FIRMWARE] = formData.firmwareVersion;
            row[DB_COLUMNS.HOURS] = formData.operatingHours;
            row[DB_COLUMNS.NOTES] = formData.techNotes;
            row[DB_COLUMNS.LAST_SERVICED_BY] = formData.email;
            row[DB_COLUMNS.VISUAL] = formData.visualImpression;
            
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