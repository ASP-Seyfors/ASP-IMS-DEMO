// ==========================================
// UPS AUTHENTICATION ENGINE
// ==========================================
function getUpsAccessToken() {
  var clientId = UPS_CONFIG.CLIENT_ID.trim();
  var clientSecret = UPS_CONFIG.CLIENT_SECRET.trim();
  
  var url = UPS_CONFIG.BASE_URL + "/security/v1/oauth/token";
  var authHeader = "Basic " + Utilities.base64Encode(clientId + ":" + clientSecret, Utilities.Charset.UTF_8);
  
  var options = {
    "method": "POST",
    "headers": {
      "Authorization": authHeader,
      "Content-Type": "application/x-www-form-urlencoded" // Explicitly required by UPS
    },
    "payload": "grant_type=client_credentials", // Hardcoded string bypasses GAS formatting bugs
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  
  if (json.access_token) {
    return json.access_token;
  } else {
    throw new Error("UPS Auth Failed: " + response.getContentText());
  }
}

// ==========================================
// SEED TEST
// ==========================================
function testUpsConnection() {
  try {
    var token = getUpsAccessToken();
    Logger.log("✅ SUCCESS! UPS is communicating.");
    Logger.log("Temporary Access Token Generated: " + token);
  } catch (err) {
    Logger.log("❌ ERROR: " + err.message);
  }
}

function testUpsHardcoded() {
  // Stitched manually from the screenshot to guarantee zero whitespace or newlines
  var clientId = "1afxK2TF7TT99PBQ6kc9AUSQVtoz544nAvkmAAnPY6EuqARK";
  var clientSecret = "HGtrnzysQoCo93aw46ivF696XN91efQbH0hYkN3qArBLI23gzgoF5JSm6ThS5oDK";
  
  var url = "https://wwwcie.ups.com/security/v1/oauth/token";
  var authHeader = "Basic " + Utilities.base64Encode(clientId + ":" + clientSecret, Utilities.Charset.UTF_8);
  
  var options = {
    "method": "POST",
    "headers": {
      "Authorization": authHeader
    },
    // Passing an object forces Apps Script to perfectly format it without charset bugs
    "payload": {
      "grant_type": "client_credentials"
    },
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  Logger.log("Raw UPS Response: " + response.getContentText());
}