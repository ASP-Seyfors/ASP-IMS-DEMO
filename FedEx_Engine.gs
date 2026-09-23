// ==========================================
// FEDEX AUTHENTICATION ENGINE
// ==========================================
function getFedExAccessToken() {
  var url = FEDEX_CONFIG.BASE_URL + "/oauth/token";
  
  // FedEx requires x-www-form-urlencoded payloads for authentication
  var payload = {
    "grant_type": "client_credentials",
    "client_id": FEDEX_CONFIG.API_KEY,
    "client_secret": FEDEX_CONFIG.SECRET_KEY
  };
  
  var options = {
    "method": "post",
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  
  if (json.access_token) {
    return json.access_token;
  } else {
    throw new Error("FedEx Auth Failed: " + response.getContentText());
  }
}

// ==========================================
// SEED TEST
// ==========================================
function testFedExConnection() {
  try {
    var token = getFedExAccessToken();
    Logger.log("✅ SUCCESS! FedEx is communicating.");
    Logger.log("Temporary Access Token Generated: " + token);
  } catch (err) {
    Logger.log("❌ ERROR connecting to FedEx: " + err.message);
  }
}

// ==========================================
// FEDEX TRACKING TEST (VIRTUAL RESPONSE)
// ==========================================
function testFedExVirtualResponse() {
  try {
    // 1. Get the 60-minute token
    var token = getFedExAccessToken();
    
    // 2. Target the FedEx Sandbox Tracking Endpoint
    var url = FEDEX_CONFIG.BASE_URL + "/track/v1/trackingnumbers";
    
    // 3. Build a standard tracking request with a dummy number
    var payload = {
      "includeDetailedScans": true,
      "trackingInfo": [
        {
          "trackingNumberInfo": {
            "trackingNumber": "123456789012"
          }
        }
      ]
    };
    
    // 4. Package the request with your Auth token
    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json",
        "X-locale": "en_US"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    // 5. Fire the request and parse the response
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    
    Logger.log("API Response Code: " + response.getResponseCode());
    
    // Look specifically for the Virtual Response alert mentioned in the docs
    if (json.output && json.output.alerts) {
      Logger.log("--- FEDEX ALERTS ---");
      json.output.alerts.forEach(alert => {
        Logger.log("Code: " + alert.code + " | Message: " + alert.message);
      });
    } else if (json.errors) {
      Logger.log("--- FEDEX ERRORS ---");
      Logger.log(JSON.stringify(json.errors, null, 2));
    }
    
  } catch (err) {
    Logger.log("❌ ERROR connecting to FedEx API: " + err.message);
  }
}

// ==========================================
// CREATE FEDEX SANDBOX SHIPMENT
// ==========================================
// ==========================================
// CREATE FEDEX SHIPMENT
// ==========================================
function createFedExShipment(payload) {
  var token = getFedExAccessToken(); 
  var endpoint = FEDEX_CONFIG.BASE_URL + "/ship/v1/shipments";
  
  // 1. Map payload data precisely from the frontend (No Sandbox fallbacks)
  var customerName = payload.customerName || "Unknown Customer";
  var contactName = payload.contactName || "Receiving Dept"; // Safe generic fallback for commercial deliveries
  
  // FedEx requires a phone number; if the customer lacks one, default to ASP's main line so the label succeeds
  var phone = String(payload.phone || FEDEX_CONFIG.SHIPPER_PHONE).replace(/\D/g, ''); 
  
  var orderNum = payload.orderNum || "N/A";
  var weight = payload.totalWeight || 1.0; 
  var serviceType = payload.serviceType || "FEDEX_GROUND";
  
  // These are guaranteed to exist because of your new frontend validation
  var street = payload.street;
  var city = payload.city;
  var state = payload.state;
  var zip = payload.zip;
  
  var requestBody = {
    "requestedShipment": {
      "shipper": {
        "contact": {
          "personName": FEDEX_CONFIG.SHIPPER_PERSON,
          "phoneNumber": FEDEX_CONFIG.SHIPPER_PHONE,
          "companyName": FEDEX_CONFIG.SHIPPER_COMPANY
        },
        "address": {
          "streetLines": [FEDEX_CONFIG.SHIPPER_STREET],
          "city": FEDEX_CONFIG.SHIPPER_CITY,
          "stateOrProvinceCode": FEDEX_CONFIG.SHIPPER_STATE,
          "postalCode": FEDEX_CONFIG.SHIPPER_ZIP,
          "countryCode": "US"
        }
      },
      "recipients": [{
        "contact": {
          "personName": contactName,
          "phoneNumber": phone,
          "companyName": customerName
        },
        "address": {
          "streetLines": [street],
          "city": city,
          "stateOrProvinceCode": state,
          "postalCode": zip,
          "countryCode": "US",
          "residential": payload.isResidential === true // ✨ REQUIRED BY FEDEX FOR RESIDENTIAL ADDRESSES
        }
      }],
      "shipDatestamp": Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd"),
      "serviceType": serviceType,
      "packagingType": "YOUR_PACKAGING",
      "pickupType": "USE_SCHEDULED_PICKUP",
      "shippingChargesPayment": {
        "paymentType": "SENDER",
        "payor": {
          "responsibleParty": {
            "accountNumber": {
              "value": FEDEX_CONFIG.ACCOUNT_NUMBER
            }
          }
        }
      },
      "labelSpecification": {
        "imageType": "PDF",
        "labelStockType": FEDEX_CONFIG.LABEL_STOCK_TYPE
      },
      "requestedPackageLineItems": [{
        "weight": {
          "units": "LB",
          "value": weight
        }
      }]
    },
    "labelResponseOptions": "LABEL", 
    "accountNumber": {
      "value": FEDEX_CONFIG.ACCOUNT_NUMBER
    }
  };

  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "x-locale": "en_US"
    },
    "payload": JSON.stringify(requestBody),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(endpoint, options);
  var json = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() >= 400 || json.errors) {
    return { status: "error", message: JSON.stringify(json.errors || json) };
  }
  
  // 2. Extract Tracking Number and Base64 PDF
  var trackingNumber = json.output.transactionShipments[0].pieceResponses[0].trackingNumber;
  var labelBase64 = json.output.transactionShipments[0].pieceResponses[0].packageDocuments[0].encodedLabel;
  
  // Format the clickable Tracking Hyperlink
  var trackLink = '=HYPERLINK("https://www.fedex.com/fedextrack/?trknbr=' + trackingNumber + '", "Track")';
  
  // 3. Write perfectly to the Outgoing tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outSheet = ss.getSheetByName("Outgoing");
  if (outSheet) {
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
    var dims = payload.dimL ? (payload.dimL + "x" + payload.dimW + "x" + payload.dimH) : "";
    var weightStr = payload.totalWeight ? payload.totalWeight + " lbs" : "";
    
    outSheet.appendRow([
      dateStr,            // A: Date
      customerName,       // B: Customer Name
      orderNum,           // C: Invoice / PO
      "FedEx",            // D: Carrier
      trackingNumber,     // E: Tracking Number
      "Yes",              // F: Tracking Email
      "Pending",          // G: Status
      "",                 // H: ETA
      "",                 // I: Final Invoice Sent?
      "Sandbox Label generated.", // J: Notes
      trackLink,          // K: Carrier Link
      "",                 // L: ETA This Week?
      "",                 // M: Delay Reason
      weightStr,          // N: Weight
      dims                // O: Dimensions
    ]);
    SpreadsheetApp.flush();
  }
  
  return { status: "success", trackingNumber: trackingNumber, label: labelBase64 };
}