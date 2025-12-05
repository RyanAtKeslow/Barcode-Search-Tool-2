// TestFormSubmit.js
// Mock event data and test functions for Camera Form Submit handlers

function testAlexa35FormSubmit() {
    const fakeEvent = {
      values: [
        "4/25/2025",                // A - SERVICE
        "testuser@keslowcamera.com",// B - EMAIL
        "9042510",              // C - BARCODE
        "62023",               // D - SERIAL
        "LPL with PL Adaptor",                       // E - MOUNT
        "",                         // F
        "",                         // G
        "",                         // H
        "5",               // I - VISUAL
        "4.0.0",                  // J - FIRMWARE
        "",                         // K
        "",                         // L
        "",                         // M
        "",                         // N
        "",                         // O
        "",                         // P
        "",                         // Q
        "Ready To Rent",            // R - STATUS
        "Testing, Sub-Rental Camera - Owen"                // S - NOTES
      ],
      range: {
        getSheet: function() {
          return {
            getName: function() {
              return "NEW Alexa 35";
            }
          };
        }
      }
    };
    CameraFormSubmit(fakeEvent);
  }
  
  function testSonyVenice2FormSubmit() {
    const fakeEvent = {
      values: [
        "5/2/2025 14:09:09",                // A - TIMESTAMP
        "veniceuser@keslowcamera.com",      // B - EMAIL
        "2029406",                          // C - BARCODE
        "10005",                            // D - SERIAL
        "123456",                           // E - SENSOR_BLOCK
        "8K",                               // F - SENSOR_RES
        "PL",                               // G - LENS_MOUNT
        "Full Service",                     // H - SERVICE_TYPE
        "OK",                               // I
        "OK",                               // J
        "No",                               // K
        "5",                                // L - VISUAL
        "V3.00",                            // M - FIRMWARE
        "1000",                             // N - HOURS
        "Full-Frame License, Anamorphic License", // O - LICENSES
        "OK",                               // P
        "OK",                               // Q
        "OK",                               // R
        "OK",                               // S
        "OK",                               // T
        "Ready To Rent",                    // U - STATUS
        "Testing Notes - Owen"                             // V - NOTES
      ],
      range: {
        getSheet: function() {
          return {
            getName: function() {
              return "NEW Sony Venice 2";
            }
          };
        }
      }
    };
    CameraFormSubmit(fakeEvent);
  } 