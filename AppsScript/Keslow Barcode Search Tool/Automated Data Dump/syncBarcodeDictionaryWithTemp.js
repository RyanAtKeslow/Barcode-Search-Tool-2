function syncBarcodeDictionaryWithTemp(ss) {
  Logger.log("🔄 [syncBarcodeDictionaryWithTemp] Starting sync of Barcode Dictionary with Temp Sheet in chunks...");
  const barcodeSheet = ss.getSheetByName('Barcode Dictionary');
  const tempSheet = ss.getSheetByName('Temp Sheet');
  if (!barcodeSheet || !tempSheet) {
    Logger.log('❌ One or both sheets not found.');
    return;
  }
  const barcodeData = barcodeSheet.getDataRange().getValues();
  const tempData = tempSheet.getDataRange().getValues();
  const maxRows = Math.max(barcodeData.length, tempData.length);
  const numCols = Math.max(barcodeData[0].length, tempData[0].length);
  const chunkSize = 200;
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 4 * 60 * 1000; // 4 minutes

  // Track progress using PropertiesService
  let startRow = Number(PropertiesService.getScriptProperties().getProperty('syncStartRow')) || 1;
  Logger.log('🔄 Starting sync at row: ' + startRow);

  // Build a map of unique keys for fast lookup
  function getKey(row) {
    return row.slice(0, 6).join('||');
  }
  const tempMap = new Map();
  for (let i = 1; i < tempData.length; i++) {
    tempMap.set(getKey(tempData[i]), tempData[i]);
  }
  const barcodeMap = new Map();
  for (let i = 1; i < barcodeData.length; i++) {
    barcodeMap.set(getKey(barcodeData[i]), {row: barcodeData[i], idx: i});
  }

  let processed = 0;
  let rowsToDelete = [];
  let rowsToAdd = [];
  // 1. Handle deletions (rows in Barcode Dictionary not in Temp Sheet)
  for (let i = startRow; i < barcodeData.length && processed < chunkSize; i++) {
    const key = getKey(barcodeData[i]);
    if (!tempMap.has(key)) {
      barcodeSheet.getRange(i+1, 1, 1, numCols).setBackground('red');
      rowsToDelete.push(i+1);
    }
    processed++;
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      Logger.log('⏳ Approaching timeout, saving progress at row ' + i);
      PropertiesService.getScriptProperties().setProperty('syncStartRow', i);
      return;
    }
  }
  // Delete from bottom up
  rowsToDelete.sort((a, b) => b - a);
  for (const rowIdx of rowsToDelete) {
    barcodeSheet.deleteRow(rowIdx);
  }

  // 2. Handle additions (rows in Temp Sheet not in Barcode Dictionary)
  processed = 0;
  for (let i = startRow; i < tempData.length && processed < chunkSize; i++) {
    const key = getKey(tempData[i]);
    if (!barcodeMap.has(key)) {
      rowsToAdd.push(tempData[i]);
    }
    processed++;
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      Logger.log('⏳ Approaching timeout, saving progress at row ' + i);
      PropertiesService.getScriptProperties().setProperty('syncStartRow', i);
      return;
    }
  }
  if (rowsToAdd.length > 0) {
    const startRowToAdd = barcodeSheet.getLastRow() + 1;
    barcodeSheet.getRange(startRowToAdd + 1, 1, rowsToAdd.length, numCols).setValues(rowsToAdd);
    barcodeSheet.getRange(startRowToAdd + 1, 1, rowsToAdd.length, numCols).setBackground('green');
  }

  // 3. Handle updates (for matching keys, update changed cells and set text color to green)
  processed = 0;
  let valuesToUpdate = [];
  let backgroundsToUpdate = [];
  let fontColorsToUpdate = [];
  let rowIndices = [];

  for (let i = startRow; i < Math.min(startRow + chunkSize, tempData.length); i++) {
    const key = getKey(tempData[i]);
    if (barcodeMap.has(key)) {
      const {row: oldRow, idx} = barcodeMap.get(key);
      for (let j = 0; j < numCols; j++) {
        if (tempData[i][j] !== oldRow[j]) {
          valuesToUpdate.push(tempData[i]);
          backgroundsToUpdate.push(new Array(numCols).fill('green'));
          fontColorsToUpdate.push(new Array(numCols).fill('green'));
          rowIndices.push(idx + 1);
        }
      }
    }
    processed++;
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      Logger.log('⏳ Approaching timeout, saving progress at row ' + i);
      PropertiesService.getScriptProperties().setProperty('syncStartRow', i);
      return;
    }
  }

  if (valuesToUpdate.length > 0) {
    barcodeSheet.getRange(rowIndices[0], 1, valuesToUpdate.length, numCols).setValues(valuesToUpdate);
    barcodeSheet.getRange(rowIndices[0], 1, backgroundsToUpdate.length, numCols).setBackgrounds(backgroundsToUpdate);
    barcodeSheet.getRange(rowIndices[0], 1, fontColorsToUpdate.length, numCols).setFontColors(fontColorsToUpdate);
  }

  // If finished, clear the progress property
  PropertiesService.getScriptProperties().deleteProperty('syncStartRow');
  Logger.log("✅ [syncBarcodeDictionaryWithTemp] Barcode Dictionary synced with Temp Sheet.");
}

function runSyncBarcodeDictionaryWithTemp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  syncBarcodeDictionaryWithTemp(ss);
} 