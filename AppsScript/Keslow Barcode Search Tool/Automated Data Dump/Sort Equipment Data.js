function sortFlawlessDataAutomationMode(inputData) {
    if (!inputData || !inputData.length) return [];

    // Define the new headers we want to use
    const headers = ["Category", "Location", "Status", "Equipment Name", "Owner", "UUID", "Barcodes"];
    
    Logger.log("🔄 Processing and concatenating barcodes...");
    const uniqueRows = new Map();
    
    // Updated column mapping for the new structure
    const COLUMNS = {
      ASSET_ID: 0,        // Column A - Asset ID (will be ignored)
      UUID: 1,            // Column B
      EQUIPMENT: 2,       // Column C - Equipment Name
      CATEGORY: 3,        // Column D - Equipment Category
      BARCODE: 4,         // Column E
      ASSET_SERIAL: 5,    // Column F - Asset Serial Number (will be ignored)
      STATUS: 6,          // Column G
      OWNER: 7,           // Column H
      LOCATION: 8         // Column I
    };
    
    // Process each row (skip header row)
    for (let i = 1; i < inputData.length; i++) {
      const row = inputData[i];
      // Handle empty barcode cells
      const barcode = row[COLUMNS.BARCODE] ? row[COLUMNS.BARCODE].toString().trim() : "No Barcode";

      // Create key from the row data (excluding barcode)
      const key = [
        row[COLUMNS.UUID],       // UUID
        row[COLUMNS.EQUIPMENT],  // Equipment Name
        row[COLUMNS.CATEGORY],   // Category
        row[COLUMNS.STATUS],     // Status
        row[COLUMNS.OWNER],      // Owner
        row[COLUMNS.LOCATION]    // Location
      ].join('|');

      if (!uniqueRows.has(key)) {
        // First occurrence of this combination
        uniqueRows.set(key, {
          uuid: row[COLUMNS.UUID],
          equipment: row[COLUMNS.EQUIPMENT],
          category: row[COLUMNS.CATEGORY],
          status: row[COLUMNS.STATUS],
          owner: row[COLUMNS.OWNER],
          location: row[COLUMNS.LOCATION] || "UNKNOWN",
          barcodes: new Set([barcode])
        });
      } else {
        // Add barcode to existing combination
        uniqueRows.get(key).barcodes.add(barcode);
      }
    }

    // Prepare final data with concatenated barcodes
    Logger.log("📝 Preparing final data...");
    const finalData = [headers];

    // Convert grouped data to rows with concatenated barcodes in new order
    for (const [key, value] of uniqueRows) {
      finalData.push([
        value.category,   // Category
        value.location,   // Location
        value.status,     // Status
        value.equipment,  // Equipment Name
        value.owner,      // Owner
        value.uuid,       // UUID
        Array.from(value.barcodes).join('|')  // Concatenated barcodes
      ]);
    }

    Logger.log(`✅ Processed ${inputData.length} rows into ${finalData.length} unique combinations`);
    return finalData;
}