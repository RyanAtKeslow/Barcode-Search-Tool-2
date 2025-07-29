function CameraFormSubmit(e) {
  var sheetName = e.range ? e.range.getSheet().getName() : null;
  if (!sheetName) {
    Logger.log('CameraFormSubmit: No sheet name found in event.');
    return;
  }

  if (sheetName === 'NEW Alexa 35') {
    Logger.log('CameraFormSubmit: Routing to alexa35FormSubmit');
    alexa35FormSubmit(e);
  } else if (sheetName === 'NEW Sony Venice 2') {
    Logger.log('CameraFormSubmit: Routing to SonyVenice2FormSubmit');
    SonyVenice2FormSubmit(e);
  } else if (sheetName === 'NEW Alexa Mini LF') {
    Logger.log('CameraFormSubmit: Routing to AMLFFormSubmit');
    AMLFFormSubmit(e);
  } else {
    Logger.log('CameraFormSubmit: No matching form handler for sheet: ' + sheetName);
  }
} 