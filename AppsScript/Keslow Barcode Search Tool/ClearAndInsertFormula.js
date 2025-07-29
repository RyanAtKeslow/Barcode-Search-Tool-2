const sheets = [
  'ER Aisle 1', 'ER Aisle 2', 'ER Aisle 3', 'ER Aisle 4', 'ER Aisle 5',
  'ER Aisle 6', 'ER Aisle 7', 'ER Aisle 8', 'ER Aisle 9', 'ER Aisle 10',
  'ER Aisle 11', 'ER Aisle 12', 'ER Aisle 13', 'ER Aisle 14',
  'Service Department', 'Battery Room', 'Filter Room',
  'Purchasing Mezzanine', 'Projector Room', 'Consignment Rooms',
  'Old Accounting'
];

function clearAndInsertFormula() {
  // First loop: Clear content from F2:F and C2:C
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      const rangeF = sheet.getRange('F2:F');
      const rangeC = sheet.getRange('C2:C');
      rangeF.clearContent();
      rangeC.clearContent();
      console.log(`Cleared content from ${sheetName} F2:F and C2:C`);
    } else {
      console.log(`Sheet ${sheetName} not found.`);
    }
  });

  // Second loop: Insert formulas into F2:F and C2:C
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      const rangeF = sheet.getRange('F2:F');
      const rangeC = sheet.getRange('C2:C');
      
      const formulaF = `=IF(C2 = "", 0, IF(C2 = "UUID Not Found", 0, IFERROR(SUMPRODUCT(--(TRIM('Barcode Dictionary'!F:F) = TRIM(C2)), LEN('Barcode Dictionary'!G:G) - LEN(SUBSTITUTE('Barcode Dictionary'!G:G, "|", "")) + 1), 0)))`;
      const formulaC = `=IF(B2="","", IFERROR(INDEX('Barcode Dictionary'!F:F, MATCH(B2, 'Barcode Dictionary'!D:D, 0)), "UUID not matched"))`;
      
      rangeF.setFormula(formulaF);
      rangeC.setFormula(formulaC);
      console.log(`Inserted formulas into ${sheetName} F2:F and C2:C`);
    } else {
      console.log(`Sheet ${sheetName} not found.`);
    }
  });
}

// Run the function
clearAndInsertFormula(); 