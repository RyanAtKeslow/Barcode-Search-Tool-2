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