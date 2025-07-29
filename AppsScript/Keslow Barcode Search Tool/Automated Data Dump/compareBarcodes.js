function compareBarcodes() {
  try {
    Logger.log("🚀 Starting barcode comparison...");
    
    // Get the active spreadsheet
    Logger.log("🔍 Getting active spreadsheet...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Could not access active spreadsheet");
    }
    
    // Get both sheets
    Logger.log("🔍 Getting sheets...");
    const importSheet = ss.getSheetByName('Barcode Dictionary Import');
    const dictSheet = ss.getSheetByName('Barcode Dictionary');
    
    if (!importSheet) {
      throw new Error("Could not find 'Barcode Dictionary Import' sheet");
    }
    if (!dictSheet) {
      throw new Error("Could not find 'Barcode Dictionary' sheet");
    }

    function extractBarcodes(value) {
      if (!value) return [];
      const stringValue = value.toString().trim();
      // Handle both single barcodes and pipe-separated lists
      return stringValue.includes('|') ? 
        stringValue.split('|').map(b => b.trim()).filter(b => b && b.toLowerCase() !== 'barcodes') :
        [stringValue].filter(b => b && b.toLowerCase() !== 'barcodes');
    }

    // Get barcodes from import sheet (Column C)
    Logger.log("📊 Reading import sheet barcodes...");
    const importLastRow = importSheet.getLastRow();
    const importBarcodes = new Set();
    
    if (importLastRow > 1) { // Start after header row
      const writeChunkSize = 50000;
      for (let i = 2; i <= importLastRow; i += writeChunkSize) {
        const endRow = Math.min(i + writeChunkSize - 1, importLastRow);
        const values = importSheet.getRange(`C${i}:C${endRow}`).getValues();
        values.flat().forEach(value => {
          if (value) {
            extractBarcodes(value).forEach(barcode => importBarcodes.add(barcode));
          }
        });
      }
    }
    Logger.log(`📊 Import sheet statistics:
    - Total rows: ${importLastRow}
    - Data rows: ${importLastRow - 1}
    - Unique barcodes: ${importBarcodes.size}`);

    // Get barcodes from dictionary sheet (Column G)
    Logger.log("📊 Reading dictionary barcodes...");
    const dictLastRow = dictSheet.getLastRow();
    const dictBarcodes = new Set();
    
    if (dictLastRow > 1) { // Start after header row
      const writeChunkSize = 50000;
      for (let i = 2; i <= dictLastRow; i += writeChunkSize) {
        const endRow = Math.min(i + writeChunkSize - 1, dictLastRow);
        const values = dictSheet.getRange(`G${i}:G${endRow}`).getValues();
        values.flat().forEach(value => {
          extractBarcodes(value).forEach(barcode => dictBarcodes.add(barcode));
        });
      }
    }
    Logger.log(`📊 Dictionary statistics:
    - Total rows: ${dictLastRow}
    - Data rows: ${dictLastRow - 1}
    - Unique barcodes: ${dictBarcodes.size}`);

    // Compare the sets to find new and missing barcodes
    const newBarcodes = [...importBarcodes].filter(x => !dictBarcodes.has(x));
    const missingFromImport = [...dictBarcodes].filter(x => !importBarcodes.has(x));

    // Create results sheet
    Logger.log("📝 Creating results sheet...");
    let resultsSheet = ss.getSheetByName('Barcode Comparison Results');
    if (resultsSheet) {
      resultsSheet.clear();
    } else {
      resultsSheet = ss.insertSheet('Barcode Comparison Results');
    }

    // Prepare summary data (ensure all rows have 2 columns)
    const summaryData = [
      ["Import Sheet Unique Barcodes:", importBarcodes.size],
      ["Dictionary Unique Barcodes:", dictBarcodes.size],
      ["New Barcodes Found:", newBarcodes.length],
      ["Missing Barcodes Count:", missingFromImport.length],
      ["", ""],
      ["New Barcodes (in Import but not in Dictionary):", ""],
      ...newBarcodes.map(b => [b, ""]),
      ["", ""],
      ["Missing Barcodes (in Dictionary but not in Import):", ""],
      ...missingFromImport.map(b => [b, ""])
    ];

    // Write results
    resultsSheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
    
    // Format the summary
    resultsSheet.getRange(1, 1, 4, 1).setFontWeight("bold");
    resultsSheet.getRange(6, 1, 1, 1).setFontWeight("bold");
    resultsSheet.getRange(8 + newBarcodes.length + 1, 1, 1, 1).setFontWeight("bold");
    resultsSheet.autoResizeColumns(1, 2);

    // Log results
    Logger.log(`📊 Comparison Results:
    - New barcodes found: ${newBarcodes.length}
    - Missing barcodes: ${missingFromImport.length}`);

    if (missingFromImport.length > 0) {
      Logger.log("⚠️ Warning: Some barcodes from dictionary are missing in the import sheet");
      // Log sample of missing barcodes
      const sampleSize = Math.min(5, missingFromImport.length);
      Logger.log(`Sample of missing barcodes: ${missingFromImport.slice(0, sampleSize).join(", ")}`);
    }

    if (newBarcodes.length > 0) {
      Logger.log("ℹ️ New barcodes found in import sheet");
      // Log sample of new barcodes
      const sampleSize = Math.min(5, newBarcodes.length);
      Logger.log(`Sample of new barcodes: ${newBarcodes.slice(0, sampleSize).join(", ")}`);
    }

    Logger.log("✅ Comparison complete!");
    
  } catch (error) {
    Logger.log(`❌ Error in compareBarcodes: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
} 