function translateCity(val) {
  if (!val) return val;
  if (val.toString().toUpperCase().includes('CULVER CITY')) return 'LA';
  if (val.toString().toUpperCase().includes('VANCOUVER')) return 'VAN';
  if (val.toString().toUpperCase().includes('ATLANTA')) return 'ATL';
  if (val.toString().toUpperCase().includes('TORONTO')) return 'TOR';
  if (val.toString().toUpperCase().includes('NEW ORLEANS')) return 'NOL';
  if (val.toString().toUpperCase().includes('ALBUQUERQUE')) return 'ABQ';
  if (val.toString().toUpperCase().includes('CHICAGO')) return 'CHI';
  return val;
}

// Translates full city names from dropdown selections to two-letter abbreviations
function translateCityDropDowns(val) {
  if (!val) return val;
  const str = val.toString().toUpperCase();
  
  // Handle full city names and convert to two-letter codes
  if (str.includes('ALBUQUERQUE') || str.includes('NEW MEXICO')) return 'ABQ';
  if (str.includes('VANCOUVER') || str.includes('BRITISH COLUMBIA')) return 'VN';
  if (str.includes('TORONTO') || str.includes('ONTARIO')) return 'TO';
  if (str.includes('LOS ANGELES') || str.includes('CULVER CITY') || str.includes('CALIFORNIA')) return 'LA';
  if (str.includes('ATLANTA') || str.includes('GEORGIA')) return 'AT';
  if (str.includes('NEW ORLEANS') || str.includes('LOUISIANA')) return 'NO';
  if (str.includes('CHICAGO') || str.includes('ILLINOIS')) return 'CH';
  
  // If no match found, return first two letters
  return str.substring(0, 2);
}

// Returns only a two-letter city code (VN, TO, LA, AT, NO, CH), with an exception for ABQ
function translateCityCode(val) {
  if (!val) return val;
  const str = val.toString().toUpperCase();
  if (str.includes('ABQ') || str.includes('ALBUQUERQUE') || str.includes('NM')) return 'ABQ';
  if (str.includes('VAN') || str.includes('BC')) return 'VN';
  if (str.includes('TOR') || str.includes('ON')) return 'TO';
  if (str.includes('LA') || str.includes('CA')) return 'LA';
  if (str.includes('ATL') || str.includes('GA')) return 'AT';
  if (str.includes('NOL') || str.includes('LA')) return 'NO';
  if (str.includes('CHI') || str.includes('IL')) return 'CH';
  return str.substring(0, 2);
} 