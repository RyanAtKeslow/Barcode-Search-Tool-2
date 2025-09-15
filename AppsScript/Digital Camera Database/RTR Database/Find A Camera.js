/**
 * Find A Camera - Camera Availability Search Tool
 * 
 * This script searches for available cameras based on multiple criteria
 * and filters results by location, date range, and group classification.
 * 
 * Step-by-step process:
 * 1. Reads filter parameters from "Look Up" sheet (camera types, location, date range, group filter)
 * 2. Expands location filters to include regional mappings (US/CAN regions)
 * 3. Identifies date columns within the specified date range
 * 4. Scans camera data starting from row 7 for matching criteria
 * 5. Checks each camera for availability (empty cells with white background)
 * 6. Extracts barcode and serial number from notes column using regex
 * 7. Classifies cameras into two groups:
 *    - Group 1 (Keslow): Matches specific barcode pattern with NBCA
 *    - Group 2 (Consigner): Contains percentage symbol in notes
 * 8. Filters results based on group selection
 * 9. Displays results in columns F-J of "Look Up" sheet
 * 
 * Filter Criteria:
 * - Camera types: Comma-separated list from A2
 * - Locations: Supports region expansion (US → multiple cities)
 * - Date range: From C2 to D2
 * - Group filter: "keslow", "consigner", or both
 * 
 * Output: Available cameras with barcode, serial, camera type, location, and notes
 */
function findAvailableCameras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Camera");
  const outputSheet = ss.getSheetByName("Look Up");

  Logger.log("Started findAvailableCameras");

  // Read filter parameters
  const [cameraRaw, locationRaw, fromDateRaw, toDateRaw, groupFilterRaw] = [
    outputSheet.getRange("A2").getValue(),
    outputSheet.getRange("B2").getValue(),
    outputSheet.getRange("C2").getValue(),
    outputSheet.getRange("D2").getValue(),
    outputSheet.getRange("E2").getValue()
  ];

  const cameraTypes = cameraRaw.toString().split(",").map(s => s.trim());
  // Build location list with region expansion and removal of empty tokens
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
  const applyLocationFilter = locations.length > 0; // true if at least one location criterion provided
  const fromDate = new Date(fromDateRaw);
  const toDate = new Date(toDateRaw);
  const groupFilter = groupFilterRaw.toString().trim().toLowerCase();

  Logger.log("Camera Types: " + cameraTypes.join(", "));
  Logger.log("Locations: " + locations.join(", "));
  Logger.log("Date Range: " + fromDate + " to " + toDate);
  Logger.log("Group Filter (E2): " + groupFilter);

  const dateHeaders = sheet.getRange("1:1").getValues()[0];
  const data = sheet.getDataRange().getValues();

  const dateColStart = 4; // Column E (0-indexed 4)
  const dateCols = [];
  for (let col = dateColStart; col < dateHeaders.length; col++) {
    const dateVal = new Date(dateHeaders[col]);
    if (!isNaN(dateVal) && dateVal >= fromDate && dateVal <= toDate) {
      dateCols.push(col);
    }
  }
  Logger.log("Date columns in range: " + dateCols.join(", "));

  if (dateCols.length === 0) {
    outputSheet.getRange(2, 6).setValue("No Results Found").setHorizontalAlignment("center");
    return;
  }

  // Updated regex for Group 1 (Keslow)
  const group1Regex = /BC#\s*\S+\s+S\/N\s*\d+.*\(NBCA\*{0,2}\)/i;

  const minCol = Math.min(...dateCols);
  const maxCol = Math.max(...dateCols);
  const width = maxCol - minCol + 1;

  const group1Results = [];
  const group2Results = [];
  let matchesFound = 0;

  for (let row = 7; row < data.length; row++) {
    const rowLoc = data[row][0];
    const rowCamera = data[row][3];
    const rowNotes = data[row][4] ? data[row][4].toString() : "";

    const cameraMatch = cameraTypes.includes(rowCamera);
    const locationMatch = applyLocationFilter ? locations.includes(rowLoc) : true;

    if (rowLoc && rowCamera && cameraMatch && locationMatch) {
      // Check the columns in date range for empty and white background
      const rangeToCheck = sheet.getRange(row + 1, minCol + 1, 1, width);
      const bgColors = rangeToCheck.getBackgrounds()[0];
      const values = rangeToCheck.getValues()[0];

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
        // Extract barcode and serial number
        const bcSnMatch = rowNotes.match(/BC#\s*([\w\-]+)\s+S\/N\s*(\d+)/i);
        const barcode = bcSnMatch ? bcSnMatch[1] : "";
        const serial = bcSnMatch ? bcSnMatch[2] : "";

        // Classification logic
        if (rowNotes.includes("%")) {
          group2Results.push([barcode, serial, rowCamera, rowLoc, rowNotes]);
          Logger.log(`Row ${row + 1} added to Group 2 (% detected)`);
          matchesFound++;
        } else if (group1Regex.test(rowNotes)) {
          group1Results.push([barcode, serial, rowCamera, rowLoc, rowNotes]);
          Logger.log(`Row ${row + 1} added to Group 1 (regex match)`);
          matchesFound++;
        } else {
          Logger.log(`Row ${row + 1} skipped — did not match Group 1 or Group 2`);
        }
      }
    }
  }

  // Decide which results to show
  let results = [];
  if (groupFilter === "keslow") {
    results = group1Results;
  } else if (groupFilter === "consigner") {
    results = group2Results;
  } else {
    results = group1Results.concat(group2Results);
  }

  // Clear previous output starting from F2, clear 5 columns wide
  const startRow = 2;
  const startCol = 6; // column F
  const maxRows = outputSheet.getMaxRows();
  outputSheet.getRange(startRow, startCol, maxRows - 1, 5).clearContent();

  if (results.length > 0) {
    const outputRange = outputSheet.getRange(startRow, startCol, results.length, 5);
    outputRange.setValues(results);
    outputRange.setHorizontalAlignment("center");
  } else {
    outputSheet.getRange(startRow, startCol).setValue("No Results Found").setHorizontalAlignment("center");
  }

  Logger.log(`Script finished. ${results.length} results returned.`);
}
