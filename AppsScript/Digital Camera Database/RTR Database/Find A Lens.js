function findAvailableLens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const criteriaSheet = ss.getSheetByName("Lens Look Up");
  const outputSheet = criteriaSheet; // Write results back to the same sheet

  Logger.log("Started findAvailableLens");

  // ---------------------------------------------------------------------------
  // READ FILTER PARAMETERS FROM A2:M2
  // ---------------------------------------------------------------------------
  // Read search parameters from row 2, columns A-M
  const searchParams = criteriaSheet.getRange("A2:M2").getValues()[0];
  
  // Extract individual parameters
  const manufacturer = searchParams[0] ? searchParams[0].toString().trim() : ""; // A2
  const series = searchParams[1] ? searchParams[1].toString().trim() : ""; // B2
  const minFocalLength = searchParams[2] ? parseFloat(searchParams[2]) : null; // C2
  const maxFocalLength = searchParams[3] ? parseFloat(searchParams[3]) : null; // D2
  const tStop = searchParams[6] ? parseFloat(searchParams[6]) : null; // G2
  const modifier1 = searchParams[7] ? searchParams[7].toString().trim() : ""; // H2
  const modifier2 = searchParams[8] ? searchParams[8].toString().trim() : ""; // I2
  const modifier3 = searchParams[9] ? searchParams[9].toString().trim() : ""; // J2
  // K2, L2, M2 are empty placeholders for future use
  
  // Build modifiers array from non-empty modifier fields
  const modifiers = [modifier1, modifier2, modifier3].filter(mod => mod !== "");
  
  // Determine focal length filter logic
  const hasMinFocal = minFocalLength !== null;
  const hasMaxFocal = maxFocalLength !== null;
  const applyFocalFilter = hasMinFocal || hasMaxFocal;
  
  // Determine if this is an exact match (only min or only max) or range (both)
  const isExactFocalMatch = (hasMinFocal && !hasMaxFocal) || (!hasMinFocal && hasMaxFocal);
  const isFocalRange = hasMinFocal && hasMaxFocal;
  
  // Create search criteria object for logging
  let focalLengthDisplay = "Any";
  if (isExactFocalMatch) {
    focalLengthDisplay = hasMinFocal ? `${minFocalLength}mm` : `${maxFocalLength}mm`;
  } else if (isFocalRange) {
    focalLengthDisplay = `${minFocalLength}-${maxFocalLength}mm`;
  }
  
  const searchCriteria = {
    manufacturer: manufacturer || "Any",
    series: series || "Any", 
    focalLength: focalLengthDisplay,
    tStop: tStop || "Any",
    modifiers: modifiers.length > 0 ? modifiers.join(", ") : "None"
  };

  // ---------------------------------------------------------------------------
  // OPTIMIZED: Create search patterns directly from A2:M2 parameters
  // ---------------------------------------------------------------------------
  const createSearchPatterns = () => {
    const patterns = [];
    
    // Manufacturer pattern (optional)
    if (manufacturer) {
      patterns.push({
        type: 'manufacturer',
        pattern: new RegExp(manufacturer.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'),
        required: true
      });
    }
    
    // Series pattern (optional)
    if (series) {
      patterns.push({
        type: 'series',
        pattern: new RegExp(series.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'),
        required: true
      });
    }
    
    // Focal length pattern (optional)
    if (applyFocalFilter) {
      let focalPattern;
      if (isExactFocalMatch) {
        const targetFocal = hasMinFocal ? minFocalLength : maxFocalLength;
        focalPattern = new RegExp(`\\b${targetFocal}\\s*mm\\b`, 'i');
      } else if (isFocalRange) {
        focalPattern = new RegExp(`\\b(\\d+(?:\\.\\d+)?)\\s*mm\\b`, 'i');
      }
      if (focalPattern) {
        patterns.push({
          type: 'focal',
          pattern: focalPattern,
          required: true,
          minFocal: minFocalLength,
          maxFocal: maxFocalLength,
          isRange: isFocalRange
        });
      }
    }
    
    // T-Stop pattern (optional)
    if (tStop) {
      patterns.push({
        type: 'stop',
        pattern: new RegExp(`(?:t|f/)\\s*${tStop}`, 'i'),
        required: true
      });
    }
    
    // Modifier patterns (all must be present)
    modifiers.forEach(modifier => {
      patterns.push({
        type: 'modifier',
        pattern: new RegExp(modifier.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'),
        required: true
      });
    });
    
    return patterns;
  };
  
  const searchPatterns = createSearchPatterns();
  
  // ---------------------------------------------------------------------------
  // OPTIMIZED: Direct pattern matching without parsing
  // ---------------------------------------------------------------------------
  const lensNameMatches = (lensName) => {
    const lensStr = lensName.toString().toLowerCase().trim();
    
    if (!lensStr) return false;
    
    // Check all patterns
    for (const patternData of searchPatterns) {
      const match = lensStr.match(patternData.pattern);
      
      if (!match && patternData.required) {
        return false; // Required pattern not found
      }
      
      // Special handling for focal length ranges
      if (patternData.type === 'focal' && patternData.isRange && match) {
        const focalValue = parseFloat(match[1]);
        if (focalValue < patternData.minFocal || focalValue > patternData.maxFocal) {
          return false; // Focal length outside range
        }
      }
    }
    
    // Handle implied manufacturer mappings
    if (manufacturer && series) {
      const manufacturerLower = manufacturer.toLowerCase();
      const seriesLower = series.toLowerCase();
      
      // Master Prime is always Zeiss
      if (seriesLower === "master prime" && manufacturerLower === "zeiss") {
        if (!lensStr.includes('zeiss') && !lensStr.includes('master prime')) {
          return false;
        }
      }
      // Add other implied manufacturer mappings as needed
    }
    
    return true; // All patterns matched
  };

  // Location – P2 (may be comma-separated list)
  const locationRaw = criteriaSheet.getRange("P2").getValue();
  // Mapping for region keywords to location lists
  const regionMap = {
    US: [
      "LOS ANGELES",
      "ATLANTA",
      "CHICAGO",
      "ALBUQUERQUE",
      "NEW ORLEANS"
    ],
    CAN: [
      "VANCOUVER",
      "TORONTO"
    ]
  };

  // Build location list by expanding any region keyword tokens
  const locations = Array.from(
    new Set(
      locationRaw
        .toString()
        .split(",")
        .map(s => s.trim())
        .filter(s => s !== "")
        .flatMap(token => {
          const upperToken = token.toUpperCase();
          if (regionMap[upperToken]) {
            return regionMap[upperToken];
          }
          return [token];
        })
    )
  );
  const applyLocationFilter = locations.length > 0; // true if criteria provided
  const locationSet = new Set(locations);

  // Date range – Q2 (from) & R2 (to)
  const fromDateRaw = criteriaSheet.getRange("Q2").getValue();
  const toDateRaw = criteriaSheet.getRange("R2").getValue();
  const fromDate = new Date(fromDateRaw);
  const toDate = new Date(toDateRaw);

  // Ownership / group filter – S2
  const groupFilter = criteriaSheet.getRange("S2").getValue().toString().trim().toLowerCase();

  // ---------------------------------------------------------------------------
  // SETUP EXCLUSION LIST – sheets we do NOT want to search
  // ---------------------------------------------------------------------------
  const excludedSheets = [
    "ESC Table of Contents",
    "Order Barcode Tool",
    "Camera",
    "Look Up",
    "Lens Look Up",
    "Wireless Follow Focus",
    "Director Viewfinders LF",
    "16mm Format"
  ];

  // ---------------------------------------------------------------------------
  // LENS SHEETS TO SCAN (numbered for efficient filtering)
  // ---------------------------------------------------------------------------
  const lensSheets = [
    "ZEISS",                    // 0
    "Leitz",                    // 1
    "Cooke",                    // 2
    "Ultimate Zoom Tab - NEW",  // 3
    "E-EF-FE-B4 Prime/Zooms",   // 4
    "Ancient Optics",           // 5
    "Large Format A-O",         // 6
    "Large Format S-Z",         // 7
    "FF Anamorphic",            // 8
    "Caldwell Chameleon SC & XC", // 9
    "OTHER (Vantage/MiniHawk/SuperBaltar/Kowa/&more)", // 10
    "SPECIALTY",                // 11
    "Laowa Lenses",             // 12
    "Anamorphic (Super 35)",    // 13
    "16mm Format",              // 14
    "Leitz Loaner - CHECK w/Zack" // 15
  ];
  
  // Manufacturer-based sheet filtering for performance
  const getSheetsToScan = (manufacturer) => {
    if (!manufacturer) return lensSheets; // No manufacturer filter = scan all sheets
    
    const manufacturerLower = manufacturer.toLowerCase();
    
    if (manufacturerLower === 'zeiss') {
      // Zeiss: Skip Leitz, Cooke, Ancient Optics, Caldwell Chameleon, Laowa, Leitz Loaner
      return lensSheets.filter((_, index) => ![1, 2, 5, 9, 12, 15].includes(index));
    }
    
    if (manufacturerLower === 'cooke') {
      // Cooke: Skip Zeiss, Leitz, Ancient Optics, Caldwell Chameleon, Laowa, Leitz Loaner
      return lensSheets.filter((_, index) => ![0, 1, 5, 9, 12, 15].includes(index));
    }
    
    if (manufacturerLower === 'leitz') {
      // Leitz: Skip Zeiss, Cooke, Ancient Optics, Caldwell Chameleon, Laowa
      return lensSheets.filter((_, index) => ![0, 2, 5, 9, 12].includes(index));
    }
    
    // For other manufacturers, scan all sheets
    return lensSheets;
  };
  
  const allowedSheetSet = new Set(getSheetsToScan(manufacturer));

  // Log search criteria after all variables are defined
  Logger.log("=== SEARCH CRITERIA ===");
  Logger.log("Manufacturer: " + searchCriteria.manufacturer);
  Logger.log("Series: " + searchCriteria.series);
  Logger.log("Focal Length: " + searchCriteria.focalLength);
  Logger.log("T-Stop: " + searchCriteria.tStop);
  Logger.log("Modifiers: " + searchCriteria.modifiers);
  Logger.log("Locations (P2): " + locations.join(", "));
  Logger.log("Date Range (Q2-R2): " + fromDate + " to " + toDate);
  Logger.log("Group Filter (S2): " + groupFilter);
  Logger.log("Sheets to scan: " + Array.from(allowedSheetSet).join(", "));

  // ---------------------------------------------------------------------------
  // OPTIMIZED: Single results array instead of separate Keslow/Consigner arrays
  // ---------------------------------------------------------------------------
  const results = [];
  let matchesFound = 0;

  // Background color for Consigner - much faster than regex
  const CONSIGNER_BG_COLOR = "#00ffff";

  // ---------------------------------------------------------------------------
  // OPTIMIZED SHEET PROCESSING
  // ---------------------------------------------------------------------------
  
  // Pre-compile regex patterns for better performance
  const bcSnRegex = /BC#\s*([\w\-]+)\s+S\/N\s*(\d+)/i;
  
  // Batch process all sheets for maximum efficiency
  const sheetsToProcess = ss.getSheets().filter(sh => allowedSheetSet.has(sh.getName()));
  
  // Process all sheets with pre-filtering by group
  for (let sheetIndex = 0; sheetIndex < sheetsToProcess.length; sheetIndex++) {
    const sheet = sheetsToProcess[sheetIndex];
    let sheetMatches = 0;

    Logger.log(`Processing sheet ${sheetIndex + 1}/${sheetsToProcess.length}: ${sheet.getName()}`);

    // Get all data at once
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length === 0) {
      Logger.log(`Skipping empty sheet: ${sheet.getName()}`);
      continue; // skip empty sheet
    }

    // Get date headers once
    const dateHeaders = data[0]; // Row 1 (0-indexed)
    
    // ---------------------------------------------------------------------------
    // OPTIMIZATION: Filter out past dates before processing data
    // ---------------------------------------------------------------------------
    const today = new Date();
    const todayString = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();
    
    // Pre-calculate date columns - only include today and future dates
    const dateColStart = 4; // Column E (0-indexed 4)
    const dateCols = [];
    const futureDateCols = [];
    
    for (let col = dateColStart; col < dateHeaders.length; col++) {
      const headerDate = new Date(dateHeaders[col]);
      if (!isNaN(headerDate)) {
        // Format header date as string for comparison
        const headerDateString = (headerDate.getMonth() + 1) + "/" + headerDate.getDate() + "/" + headerDate.getFullYear();
        
        // Only include dates that are today or in the future
        if (headerDateString >= todayString && headerDate >= fromDate && headerDate <= toDate) {
          dateCols.push(col);
          futureDateCols.push(col);
        } else if (headerDate >= fromDate && headerDate <= toDate) {
          // Include dates within search range even if in past (for historical searches)
          dateCols.push(col);
        }
      }
    }
    
    if (dateCols.length === 0) {
      Logger.log("No date columns within range on sheet: " + sheet.getName());
      continue;
    }

    // Log optimization info
    if (futureDateCols.length < dateCols.length) {
      Logger.log(`Sheet ${sheet.getName()}: Filtered out ${dateCols.length - futureDateCols.length} past date columns`);
    }

    const minCol = Math.min(...dateCols);
    const maxCol = Math.max(...dateCols);
    const width = maxCol - minCol + 1;

    // ---------------------------------------------------------------------------
    // OPTIMIZATION: Batch API calls to reduce simultaneous invocations
    // ---------------------------------------------------------------------------
    try {
      // Get background colors for column E to pre-filter by group
      const notesBgColors = sheet.getRange(8, 5, data.length - 7, 1).getBackgrounds(); // Column E
      
      // Add small delay to prevent API rate limiting
      Utilities.sleep(100);
      
      // Get date range data for availability checking in one batch
      const dateRangeData = sheet.getRange(8, minCol + 1, data.length - 7, width);
      const allBgColors = dateRangeData.getBackgrounds();
      const allValues = dateRangeData.getValues();
      
      // Add small delay to prevent API rate limiting
      Utilities.sleep(100);
      
      // Pre-filter rows by group based on search criteria
      const filteredRows = [];
      
      for (let row = 7; row < data.length; row++) {
        const rowLoc = data[row][0] ? data[row][0].toString().trim() : ""; // Column A
        const rowCameraRaw = data[row][3]; // Column D
        const rowCamera = rowCameraRaw ? rowCameraRaw.toString().trim() : ""; // Trimmed value
        const rowNotes = data[row][4] ? data[row][4].toString() : ""; // Column E

        // Skip rows without camera name or location
        if (!rowCamera || rowLoc === "") continue;
        
        // Check location filter early
        const locationMatch = applyLocationFilter ? locationSet.has(rowLoc) : true;
        if (!locationMatch) continue;

        // Determine group based on background color
        const rowIdx = row - 7;
        const notesBgColor = notesBgColors[rowIdx][0].toLowerCase();
        const isConsigner = notesBgColor === CONSIGNER_BG_COLOR.toLowerCase();
        
        // Pre-filter by group based on search criteria
        if (groupFilter === "keslow" && isConsigner) continue;
        if (groupFilter === "consigner" && !isConsigner) continue;
        
        filteredRows.push({ row, rowLoc, rowCamera, rowNotes, rowIdx, isConsigner });
      }
      
            Logger.log(`Sheet ${sheet.getName()} pre-filtered: ${filteredRows.length} rows`);

      // Process filtered rows
      for (const rowData of filteredRows) {
        const cameraMatch = lensNameMatches(rowData.rowCamera);
        if (!cameraMatch) {
          // Debug: Log first few non-matching lens names
          if (sheetMatches < 3) {
            Logger.log(`Non-match: "${rowData.rowCamera}" in ${sheet.getName()}`);
          }
          continue;
        }

        // Check availability
        const bgColors = allBgColors[rowData.rowIdx];
        const values = allValues[rowData.rowIdx];

        let allEmptyAndWhite = true;
        for (let i = 0; i < dateCols.length; i++) {
          const relIdx = dateCols[i] - minCol;
          if (values[relIdx] !== "" && values[relIdx] !== null) {
            allEmptyAndWhite = false;
            break;
          }
          const bg = bgColors[relIdx].toLowerCase();
          if (bg !== "#ffffff" && bg !== "#fff" && bg !== "white") {
            allEmptyAndWhite = false;
            break;
          }
        }

        if (allEmptyAndWhite) {
          const bcSnMatch = rowData.rowNotes.match(bcSnRegex);
          const barcode = bcSnMatch ? bcSnMatch[1] : "";
          const serial = bcSnMatch ? bcSnMatch[2] : "";
          
          results.push([barcode, serial, rowData.rowCamera, rowData.rowLoc, rowData.rowNotes]);
          sheetMatches++;
          matchesFound++;
        }
      }

      // After processing all rows in current sheet, output a summary log
      if (sheetMatches > 0) {
        Logger.log(`Sheet ${sheet.getName()} summary: ${sheetMatches} matches found`);
      } else if (filteredRows.length > 0) {
        Logger.log(`Sheet ${sheet.getName()}: No lens name matches found (${filteredRows.length} rows processed)`);
      }
      
    } catch (error) {
      Logger.log(`Error processing sheet ${sheet.getName()}: ${error.message}`);
      continue; // Skip this sheet and continue with the next one
    }
  }

  // ---------------------------------------------------------------------------
  // OUTPUT – clear previous content and write new results
  // We'll start at column T (20) row 2, five columns wide (barcode, serial, camera, loc, notes)
  // ---------------------------------------------------------------------------
  const startRow = 2;
  const startCol = 20; // column T
  const maxRows = outputSheet.getMaxRows();
  outputSheet.getRange(startRow, startCol, maxRows - 1, 5).clearContent();

  if (results.length > 0) {
    const outputRange = outputSheet.getRange(startRow, startCol, results.length, 5);
    outputRange.setValues(results);
    outputRange.setHorizontalAlignment("center");
  } else {
    outputSheet
      .getRange(startRow, startCol)
      .setValue("No Results Found")
      .setHorizontalAlignment("center");
  }

  Logger.log(`Script finished. ${results.length} results returned across ${matchesFound} matching rows.`);
} 