// ==========================================
// QUICKBOOKS API ENGINE
// ==========================================
function getQboAccessToken() {
  var tokenUrl = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
  var authHeader = "Basic " + Utilities.base64Encode(QBO_CONFIG.CLIENT_ID + ":" + QBO_CONFIG.CLIENT_SECRET);
  
  var payload = {
    grant_type: "refresh_token",
    refresh_token: QBO_CONFIG.REFRESH_TOKEN
  };
  
  var options = {
    method: "post",
    headers: {
      "Authorization": authHeader,
      "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: payload,
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(tokenUrl, options);
  var json = JSON.parse(response.getContentText());
  
  if (json.access_token) {
    // If Intuit returns a rotated refresh token, update script properties
    if (json.refresh_token) {
      QBO_CONFIG.REFRESH_TOKEN = json.refresh_token;
      PropertiesService.getScriptProperties().setProperty("QBO_REFRESH_TOKEN", json.refresh_token);
    }
    return json.access_token;
  } else {
    throw new Error("Failed to refresh QBO Token: " + response.getContentText());
  }
}

function fetchOpenInvoicesFromQBO() {
  var savedToken = PropertiesService.getScriptProperties().getProperty("QBO_REFRESH_TOKEN");
  if (savedToken) QBO_CONFIG.REFRESH_TOKEN = savedToken;

  var accessToken = getQboAccessToken();
  var baseUrl = QBO_CONFIG.ENVIRONMENT === "sandbox" 
    ? "https://sandbox-quickbooks.api.intuit.com" 
    : "https://quickbooks.api.intuit.com";
    
  var query = encodeURIComponent("select * from Invoice where Balance > '0'");
  var queryUrl = baseUrl + "/v3/company/" + QBO_CONFIG.REALM_ID + "/query?query=" + query + "&minorversion=65";
  
  var options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(queryUrl, options);
  var resJson = JSON.parse(response.getContentText());
  
  if (!resJson.QueryResponse || !resJson.QueryResponse.Invoice) {
    return { status: "success", count: 0, message: "No open invoices found in QBO." };
  }
  
  var invoices = resJson.QueryResponse.Invoice;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feedSheet = ss.getSheetByName("QBO_Feed");
  if (!feedSheet) {
    feedSheet = ss.insertSheet("QBO_Feed");
    feedSheet.appendRow(["Timestamp", "Session Name", "Customer", "PO / Order #", "REF / SKU", "Quantity", "Status"]);
  }
  
  var feedData = feedSheet.getDataRange().getValues();
  var existingKeys = new Set();
  
  // ✨ FIX 1: Look at Index 1 (Session Name) and Index 4 (REF/SKU)
  for (var i = 1; i < feedData.length; i++) {
    var sName = String(feedData[i][1] || '').trim().toUpperCase(); 
    var sSku = String(feedData[i][4] || '').trim().toUpperCase();  
    if (sName && sSku) existingKeys.add(sName + "_" + sSku);
  }
  
  var rowsToAdd = [];

  invoices.forEach(function(inv) {
    var invNo = inv.DocNumber || inv.Id;
    var sessionKey = "QBO Invoice #" + invNo;
    var customerName = (inv.CustomerRef && inv.CustomerRef.name) ? inv.CustomerRef.name : "UNKNOWN";
    var lines = inv.Line || [];
    
    lines.forEach(function(line) {
      if (line.DetailType === "SalesItemLineDetail" && line.SalesItemLineDetail) {
        var itemDetail = line.SalesItemLineDetail;
        var sku = (itemDetail.ItemRef && itemDetail.ItemRef.name) ? itemDetail.ItemRef.name.toUpperCase().trim() : "";
        var qty = parseInt(itemDetail.Qty, 10) || 1;
        
        if (sku) {
          var checkKey = sessionKey.toUpperCase() + "_" + sku;
          if (!existingKeys.has(checkKey)) {
            var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
            rowsToAdd.push([timestamp, sessionKey, customerName, "INV#" + invNo, sku, qty, "PENDING"]);
            existingKeys.add(checkKey);
          }
        }
      }
    });
  });
  
  if (rowsToAdd.length > 0) {
    feedSheet.getRange(feedSheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
  }

  // ✨ FIX 2: Added the logic to count ONLY pending invoices
  var pendingInvoices = new Set();
  for (var k = 1; k < feedData.length; k++) {
      var status = String(feedData[k][6] || '').toUpperCase().trim();
      if (status !== 'COMPLETED') {
          pendingInvoices.add(String(feedData[k][1] || '').trim().toUpperCase());
      }
  }
  rowsToAdd.forEach(function(r) {
      pendingInvoices.add(String(r[1] || '').trim().toUpperCase());
  });
  
  return { status: "success", count: pendingInvoices.size, addedLines: rowsToAdd.length };
}

function writeBackToQboInvoice(sessionObj) {
  if (!sessionObj || !sessionObj.orderNum) {
    return { status: "error", message: "Missing order number in payload." };
  }
  
  // Extract just the numerical invoice ID (e.g., "INV#1045" -> "1045")
  var invoiceId = String(sessionObj.orderNum).replace(/\D/g, '');
  if (!invoiceId) {
    return { status: "error", message: "Could not parse numerical invoice ID from: " + sessionObj.orderNum };
  }

  // Authenticate
  var savedToken = PropertiesService.getScriptProperties().getProperty("QBO_REFRESH_TOKEN");
  if (savedToken) QBO_CONFIG.REFRESH_TOKEN = savedToken;
  var accessToken = getQboAccessToken();
  
  var baseUrl = QBO_CONFIG.ENVIRONMENT === "sandbox" 
    ? "https://sandbox-quickbooks.api.intuit.com" 
    : "https://quickbooks.api.intuit.com";
  
  // ==========================================
  // 1. GET THE EXISTING INVOICE FROM QBO
  // ==========================================
  var getUrl = baseUrl + "/v3/company/" + QBO_CONFIG.REALM_ID + "/invoice/" + invoiceId + "?minorversion=65";
  var getOptions = {
    method: "get",
    headers: { "Authorization": "Bearer " + accessToken, "Accept": "application/json" },
    muteHttpExceptions: true
  };
  
  var getResponse = UrlFetchApp.fetch(getUrl, getOptions);
  var getJson = JSON.parse(getResponse.getContentText());
  
  if (!getJson.Invoice) {
    return { status: "error", message: "Invoice #" + invoiceId + " not found in QBO." };
  }
  
  var invoice = getJson.Invoice;
  var isModified = false;
  var scannedItems = sessionObj.scannedObjects || [];
  
  // ==========================================
  // 2. MODIFY THE LINE ITEMS
  // ==========================================
  if (invoice.Line && invoice.Line.length > 0) {
    invoice.Line.forEach(function(line) {
      if (line.DetailType === "SalesItemLineDetail" && line.SalesItemLineDetail && line.SalesItemLineDetail.ItemRef) {
        // QBO usually stores the SKU/REF in the ItemRef.name property
        var lineSku = String(line.SalesItemLineDetail.ItemRef.name || "").toUpperCase().trim();
        
        // Find all scanned instances of this SKU from the session payload
        var matchingScans = scannedItems.filter(function(scan) {
          return String(scan.ref).toUpperCase().trim() === lineSku;
        });
        
        if (matchingScans.length > 0) {
          var trackingStrings = [];
          matchingScans.forEach(function(scan) {
             // Only append if there is actually a Lot number recorded
             if (scan.lot && scan.lot !== 'NO_LOT') {
                var expStr = (scan.exp && scan.exp !== 'NO_EXP') ? " | Exp: " + scan.exp : "";
                trackingStrings.push("[Lot: " + scan.lot + expStr + "]");
             }
          });
          
          // If we found valid Lot/Exp data, append it to the QBO line item description
          if (trackingStrings.length > 0) {
            var appendStr = trackingStrings.join(" ");
            line.Description = (line.Description ? line.Description + "\n" : "") + appendStr;
            isModified = true;
          }
        }
      }
    });
  }
  
  if (!isModified) {
    return { status: "success", message: "No Lot/Exp data present to write back for Invoice #" + invoiceId };
  }
  
  // ==========================================
  // 3. POST THE UPDATED INVOICE BACK TO QBO
  // ==========================================
  var postUrl = baseUrl + "/v3/company/" + QBO_CONFIG.REALM_ID + "/invoice?minorversion=65";
  var postOptions = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Accept": "application/json",
      "Content-Type": "application/json"
    },
    // We pass the entire modified invoice object back to QBO
    payload: JSON.stringify(invoice), 
    muteHttpExceptions: true
  };
  
  var postResponse = UrlFetchApp.fetch(postUrl, postOptions);
  var postJson = JSON.parse(postResponse.getContentText());
  
  if (postJson.Invoice) {
    return { status: "success", message: "Successfully updated QBO Invoice #" + invoiceId + " with Lot/Exp data." };
  } else {
    return { status: "error", message: "Failed to update QBO Invoice.", details: postResponse.getContentText() };
  }
}