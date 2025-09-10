/**
 * Sort Equipment Data - Equipment Data Processing and Consolidation Script
 * 
 * This script processes equipment data by grouping rows with identical metadata
 * and consolidating their barcodes into pipe-delimited strings for automation.
 * 
 * Step-by-step process:
 * 1. Receives input data array and validates it's not empty
 * 2. Defines target headers: Category, Location, Status, Equipment Name, Owner, UUID, Barcodes
 * 3. Maps source columns to target structure (ignores Asset ID and Asset Serial)
 * 4. Groups rows by unique combination of metadata (UUID, Equipment, Category, Status, Owner, Location)
 * 5. Collects all barcodes for each unique combination using Set
 * 6. Handles empty barcodes by setting "No Barcode" placeholder
 * 7. Concatenates barcodes using pipe (|) separator
 * 8. Returns processed data with consolidated barcodes
 * 
 * Data Processing:
 * - Grouping: Uses Map with composite keys for efficient grouping
 * - Barcode collection: Uses Set to avoid duplicates within groups
 * - Concatenation: Joins barcodes with pipe separator
 * - Empty handling: Replaces empty barcodes with "No Barcode"
 * 
 * Column Mapping:
 * - Source: Asset ID, UUID, Equipment, Category, Barcode, Asset Serial, Status, Owner, Location
 * - Target: Category, Location, Status, Equipment Name, Owner, UUID, Barcodes
 * - Ignored: Asset ID, Asset Serial (not needed in final output)
 * 
 * Output Format:
 * - Headers: Category, Location, Status, Equipment Name, Owner, UUID, Barcodes
 * - Barcodes: Pipe-delimited string (e.g., "BC001|BC002|BC003")
 * - Empty barcodes: Replaced with "No Barcode" placeholder
 * 
 * Features:
 * - Data consolidation and deduplication
 * - Flexible column mapping
 * - Empty value handling
 * - Efficient grouping algorithm
 * - Automation-ready output format
 */
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