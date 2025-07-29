function manualConcatenateBarcodes() {
  try {
    Logger.log("🚀 Starting barcode concatenation process...");
    
    // 1. Get the spreadsheet and source sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('Barcode Dictionary Import');
    if (!sourceSheet) throw new Error("Could not find 'Barcode Dictionary Import' sheet");
    
    // 2. Read data starting from row 2
    Logger.log("📥 Reading data from row 2...");
    const lastRow = sourceSheet.getLastRow();
    const data = sourceSheet.getRange(2, 1, lastRow - 1, 7).getValues(); // Get A2:G[last]
    Logger.log(`📊 Read ${data.length} rows of data`);

    // 3. Group and concatenate barcodes
    Logger.log("🔄 Grouping and concatenating barcodes...");
    const uniqueRows = new Map();
    
    // Process each row and group by unique combination
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const barcode = row[2]; // Barcode is in column C (index 2)
      if (!barcode) continue; // Skip empty barcodes

      // Create key from the row data (excluding barcode)
      const key = [
        row[0],  // UUID
        row[1],  // Equipment Name
        row[3],  // Category
        row[4],  // Status
        row[5],  // Owner
        row[6]   // Location
      ].join('|');

      if (!uniqueRows.has(key)) {
        // First occurrence of this combination
        uniqueRows.set(key, {
          uuid: row[0],
          equipment: row[1],
          category: row[3],
          status: row[4],
          owner: row[5],
          location: row[6],
          barcodes: new Set([barcode.toString().trim()])
        });
      } else {
        // Add barcode to existing combination
        uniqueRows.get(key).barcodes.add(barcode.toString().trim());
      }
    }

    // 4. Prepare final data with new column order and concatenated barcodes
    Logger.log("📝 Preparing final data...");
    const newHeaders = ["Category", "Location", "Status", "Equipment Name", "Owner", "UUID", "Barcodes"];
    const finalData = [newHeaders];

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

    // 5. Create output sheet
    Logger.log("📄 Setting up output sheet...");
    let outputSheet = ss.getSheetByName('Concatenated Barcodes');
    if (outputSheet) {
      outputSheet.clear();
    } else {
      outputSheet = ss.insertSheet('Concatenated Barcodes');
    }

    // 6. Write the processed data in chunks
    Logger.log("📤 Writing processed data to new sheet...");
    const writeChunkSize = 5000;
    for (let i = 0; i < finalData.length; i += writeChunkSize) {
      const chunk = finalData.slice(i, Math.min(i + writeChunkSize, finalData.length));
      outputSheet.getRange(i + 1, 1, chunk.length, chunk[0].length).setValues(chunk);
      Logger.log(`📝 Wrote chunk of ${chunk.length} rows (rows ${i + 1} to ${i + chunk.length})`);
    }

    // 7. Delete columns H through Z if they exist
    Logger.log("🗑️ Deleting excess columns...");
    const lastColumn = outputSheet.getLastColumn();
    if (lastColumn > 7) { // If there are columns after G
      outputSheet.deleteColumns(8, lastColumn - 7); // Delete from H to last column
    }

    Logger.log(`✅ Complete! Processed ${data.length} rows into ${finalData.length} unique combinations`);
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.toString()}`);
    throw error;
  }
}