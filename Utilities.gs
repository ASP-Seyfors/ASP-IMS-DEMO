/*function oneTimePlatformSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Items");
  var data = sheet.getDataRange().getValues();
  
  // 1. Dynamically find the exact column indexes so we never overwrite the wrong column
  var headers = data[0];
  var revMedCol = -1, thriveCol = -1, shopifyCol = -1;
  
  for (var c = 0; c < headers.length; c++) {
    var hName = String(headers[c]).toUpperCase();
    if (hName.indexOf("REVMED") !== -1 && hName.indexOf("PRICE") === -1) revMedCol = c + 1;
    if (hName.indexOf("THRIVE") !== -1) thriveCol = c + 1;
    if (hName.indexOf("SHOPIFY") !== -1) shopifyCol = c + 1;
  }
  
  if (thriveCol === -1) {
    SpreadsheetApp.getUi().alert("Could not find the 'Synced with Thrive' column!");
    return;
  }

  // 2. The EXACT arrays from your provided files
  var revmedRefs = ["1200710","309647","3311-8G","3350-8G","134051","EGIA45AV","EGIA45AVM","LF1212A","LF1837","LF1937","LF4418","LF5637","LS1020","SIG45AMT","SIG45AVM","SIG45AXT","SIG45CTAMT","SIG45CTAVM","SIG60AMT","SIG60AVM","SIG60AXT","SIG60CTAVM","TA30V3L","1962","1963","1973","1974","1611G","1629H","1647G","1661G","1665G","1667G","1669H","1698G","1761G","1780G","1797G","1893G","1951S","1952S","1953S","2815G","3846T","424H","4350XL","460T","585H","627H","628H","632G","635H","640G","662H","663G","663H","664H","669H","678G","680H","682G","685G","687G","689G","697G","698G","698H","699G","699H","736G","770G","7731G","774G","783G","790G","792G","810H","811H","825G","8424H","8521H","8556H","8581H","8661G","8661H","8665G","8681G","8682G","8683G","8684G","8689H","8690H","8695G","8696G","8697G","8698G","8776H","882H","8832H","8871H","J206G","J212H","J214H","J219H","J258H","J259H","J269H","J304H","J315H","J316H","J328H","J338H","J340H","J344H","J345H","J347H","J385H","J415H","J417H","J463G","J464G","J488G","J489G","J492G","J492H","J493G","J494G","J495G","J496G","J496H","J497G","J497H","J506G","J510G","J544G","J546G","J548G","J566G","J570G","J571G","J576G","J578G","J599G","J833G","J868H","J906G","J944H","J945H","J946H","J947H","J975G","K802H","K832H","K833H","K834H","K891H","L880G","MB66G","RR750","84000611"];
  
  var shopifyExcludedRefs = ["AR-9821","383512","2030","EGIA45AMT","EGIA60AMT","EGIA60AVM","LF1923","SIG60CTAMT","1740G","1856G","1915G","686G","794G","796G","843H","8706H","8726G","C521D","ECR45M","ECR60M","G123H","G697G","GST45D","GST45G","GST45T","HAR9F","J195H","J346H","J493H","J562G","J839D","J946G","PMW35","V549G","VCP213H","VCP304H","VCP351H","VCP494G","VCP496H","VCP497H","VCP527H","VCP682H","VCP833G","VCP834G","VCP936H","YY31G","Z494G","202050","480455","48230B","48230W","Y303D"];
  
  var thriveRefs = ["1200710", "134051", "1611G", "1629H", "1647G", "1661G", "1665G", "1667G", "1669H", "1698G", "1740G", "1761G", "1780G", "1797G", "1856G", "1893G", "1915G", "1951S", "1952S", "1953S", "1962", "1963", "1973", "1974", "202050", "2030", "206520", "2815G", "309647", "33-31", "3311-8G", "3350-8G", "383512", "3846T", "424H", "4301-02", "4350XL", "4406", "460T", "470007", "470194", "470230", "471344", "480422", "480455", "48230B", "48230W", "48345B", "48345G", "48345M", "48345T", "48345W", "48360B", "48360G", "48360M", "48360T", "48360W", "5086-02", "585H", "627H", "628H", "632G", "635H", "640G", "662H", "663G", "663H", "664H", "669H", "678G", "680H", "682G", "685G", "686G", "687G", "689G", "697G", "698G", "698H", "699G", "699H", "71111579", "736G", "7510100", "7510200", "7510400", "7510600", "7510800", "770G", "7731G", "774G", "783G", "790G", "792G", "794G", "810H", "811H", "825G", "84000611", "8424H", "843H", "8521H", "8556H", "8557H", "8581H", "8661G", "8661H", "8665G", "8681G", "8682G", "8683G", "8684G", "8689H", "8690H", "8695G", "8696G", "8697G", "8698G", "8706H", "8726G", "8776H", "882H", "8832H", "8871H", "A1667N", "AR-7200", "AR-7202", "AR-7211", "AR-9821", "B1915N", "C521D", "CB5LT", "DHVM12", "DNX12", "ECR45M", "ECR60M", "EGIA45AMT", "EGIA45AV", "EGIA45AVM", "EGIA60AMT", "EGIA60AVM", "G121H", "G122H", "G181H", "G316N", "G346N", "GST45B", "GST45D", "GST45G", "GST45T", "GST60W", "HAR9F", "J206G", "J212H", "J214H", "J219H", "J259H", "J269H", "J304H", "J315H", "J316H", "J328H", "J338H", "J340H", "J344H", "J345H", "J347H", "J385H", "J415H", "J417H", "J463G", "J464G", "J489G", "J492G", "J492H", "J493G", "J494G", "J495G", "J496G", "J496H", "J497G", "J497H", "J506G", "J510G", "J544G", "J546G", "J548G", "J566G", "J570G", "J571G", "J576G", "J578G", "J599G", "J833G", "J839D", "J868H", "J906G", "J944H", "J945H", "J946G", "J946H", "J947H", "J975G", "K802H", "K832H", "K833H", "K834H", "K891H", "L880G", "LF1212A", "LF1837", "LF1923", "LF1937", "LF4418", "LF5637", "LS1020", "M8702", "MB66G", "MCP316H", "MCP317H", "MCP427H", "MCP493G", "MCP493H", "MCP494G", "MCP494H", "MCP496G", "MCP496H", "MCP523H", "MCP936H", "MR8-10MH17", "MR8-10MH30", "MR8-14MH30", "MX140", "PDP311H", "PDP316H", "PDP317H", "PDP334H", "PDP338H", "PDP340H", "PDP467H", "PDPB9916", "PXW35", "PXX22N", "PXX82N", "R340N", "R647H", "RR750", "SA6H", "SIG45AMT", "SIG45AVM", "SIG45AXT", "SIG45CTAMT", "SIG45CTAVM", "SIG60AMT", "SIG60AVM", "SIG60AXT", "SIG60CTAMT", "SIG60CTAVM", "TA30V3L", "V549G", "VCP213H", "VCP214H", "VCP267H", "VCP305H", "VCP311H", "VCP316H", "VCP317H", "VCP329H", "VCP332H", "VCP334H", "VCP340H", "VCP341H", "VCP356H", "VCP357H", "VCP415H", "VCP417H", "VCP421H", "VCP422H", "VCP423H", "VCP433H", "VCP441H", "VCP442H", "VCP443H", "VCP493G", "VCP504G", "VCP596H", "VCP603H", "VCP880T", "VR416", "VR493", "VR494", "VR834", "VR845", "VR945", "W31G", "X916H", "Y213H", "Y215H", "Y266H", "Y303D", "Y303H", "Y304H", "Y305H", "Y315H", "Y317H", "Y318H", "Y333H", "Y334H", "Y346H", "Y415H", "Y415N", "Y416H", "Y417H", "Y427H", "Y463G", "Y464G", "Y489G", "Y489H", "Y490H", "Y492G", "Y492H", "Y493G", "Y493N", "Y494G", "Y496G", "Y496H", "Y497G", "Y497H", "Y527H", "Y682H", "Y936H", "YY31G", "Z195T", "Z304H", "Z305H", "Z311H", "Z316H", "Z317H", "Z332H", "Z333H", "Z334H", "Z338H", "Z340H", "Z341H", "Z346H", "Z357H", "Z358T", "Z371T", "Z423H", "Z432H", "Z443H", "Z463G", "Z466H", "Z487G", "Z493G", "Z496G", "Z497G", "Z513G", "Z968H", "Z970H"];

  for (var i = 1; i < data.length; i++) {
    var ref = String(data[i][0]).toUpperCase().trim();
    if (!ref) continue;
    
    // 3. Logic Rules based on exports (USING BOOLEANS so checkboxes don't break!)
    var onRevMed = revmedRefs.indexOf(ref) !== -1 ? true : false;
    var onShopify = shopifyExcludedRefs.indexOf(ref) === -1 ? true : false;
    var onThrive = thriveRefs.indexOf(ref) !== -1 ? true : false; 
    
    sheet.getRange(i + 1, revMedCol).setValue(onRevMed);
    sheet.getRange(i + 1, thriveCol).setValue(onThrive);
    sheet.getRange(i + 1, shopifyCol).setValue(onShopify);
  }
  
  SpreadsheetApp.flush();
  
  // Optional: Pop an alert in Google Sheets so you know it finished
  var ui = SpreadsheetApp.getUi();
  if (ui) ui.alert("Database Flags successfully synced with Thrive Export data!");
}*/

// --- NEW: COLD STORAGE ARCHIVER ---
function runMonthlyColdStorageArchive() {
  var prodSs = SpreadsheetApp.getActiveSpreadsheet();
  var archiveId = PropertiesService.getScriptProperties().getProperty("ARCHIVE_SHEET_ID");
  if (!archiveId) { Logger.log("Error: ARCHIVE_SHEET_ID not found."); return; }
  
  var coldSs = SpreadsheetApp.openById(archiveId);
  var today = new Date();
  var cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 60); // 60 Days retention
  
  var sheetsToProcess = [ 
    { name: "Archive", dateCol: 0 }, 
    { name: SHEET_AUDIT, dateCol: 0 },
    { name: "QBO_Feed", dateCol: 0 } // <-- Added!
  ];

  sheetsToProcess.forEach(function(sheetInfo) {
    var pSheet = prodSs.getSheetByName(sheetInfo.name);
    if (!pSheet) return;
    
    var cSheet = coldSs.getSheetByName(sheetInfo.name);
    if (!cSheet) {
        cSheet = coldSs.insertSheet(sheetInfo.name);
        var headers = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues();
        cSheet.appendRow(headers[0]);
    }
    
    // Split into new tab if exceeding 100k rows
    if (cSheet.getLastRow() > 100000) {
      var tabCount = 2;
      while (coldSs.getSheetByName(sheetInfo.name + "_Part" + tabCount)) { tabCount++; }
      cSheet = coldSs.insertSheet(sheetInfo.name + "_Part" + tabCount);
      var headers = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues();
      cSheet.appendRow(headers[0]);
    }
    
    var data = pSheet.getDataRange().getValues();
    var rowsToArchive = [];
    var rowsToDelete = [];
    
    for (var i = data.length - 1; i > 0; i--) {
      var rowDate = new Date(data[i][sheetInfo.dateCol]);
      if (rowDate < cutoffDate) {
        rowsToArchive.push(data[i]);
        rowsToDelete.push(i + 1); 
      }
    }
    
    if (rowsToArchive.length > 0) {
      rowsToArchive.reverse();
      cSheet.getRange(cSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
      rowsToDelete.forEach(function(rowIdx) { pSheet.deleteRow(rowIdx); });
    }
  });
}

// --- NEW: AUTOMATED GITHUB COMMITS ---
function commitDatabaseJsonToGitHub() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = { items: [], customers: [], suppliers: [], vendors: [], users: [] };
  
  var userSheet = ss.getSheetByName("Users") || ss.getSheetByName("Users_List");
  if (userSheet) {
      var userData = userSheet.getDataRange().getValues();
      for (var i = 1; i < userData.length; i++) { if (userData[i][0]) db.users.push(String(userData[i][0]).trim()); }
  } else { db.users = ["Trey", "Thomas", "Jessica", "+ New User"]; }
  
  var custSheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (custSheet) { var custData = custSheet.getDataRange().getValues(); for (var i = 1; i < custData.length; i++) { if (custData[i][0]) db.customers.push(String(custData[i][0]).trim()); } }
  var supSheet = ss.getSheetByName(SHEET_SUPPLIERS);
  if (supSheet) { var supData = supSheet.getDataRange().getValues(); for (var i = 1; i < supData.length; i++) { if (supData[i][0]) db.suppliers.push(String(supData[i][0]).trim()); } }
  var mfrSheet = ss.getSheetByName(SHEET_MFRS);
  if (mfrSheet) { var mfrData = mfrSheet.getDataRange().getValues(); for (var i = 1; i < mfrData.length; i++) { if (mfrData[i][0]) db.vendors.push(String(mfrData[i][0]).trim()); } }

  var itemsSheet = ss.getSheetByName(SHEET_ITEMS);
  if (itemsSheet) {
    var itemsData = itemsSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < itemsData.length; i++) {
      var row = itemsData[i];
      if (row[0]) db.items.push({ 
        ref: String(row[0]), mfr: String(row[1]), desc: String(row[2]), gtin: String(row[3]), price: String(row[4]),
        cost: String(row[5] || "$0.00"), onHand: parseInt(row[6], 10) || 0, reservedQty: parseInt(row[7], 10) || 0,    
        availableQty: parseInt(row[8], 10) || 0, onRevMed: String(row[9] || "FALSE"), revMedPrice: String(row[10] || ""),        
        syncedShopify: String(row[13] || "FALSE"), category: String(row[14] || ""),           
        status: String(row[15] || "INACTIVE"), shelf: String(row[18] || ""),
        shopifyCategory: String(row[19] || "Medical Supplies")
      });
    }
  }

  var uomSheet = ss.getSheetByName(SHEET_BUNDLES);
  if (uomSheet) {
    var uomData = uomSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < uomData.length; i++) {
      var row = uomData[i];
      if (row[0]) db.items.push({ 
        ref: String(row[0]), mfr: String(row[1]), desc: String(row[2]), gtin: String(row[3]), price: String(row[4]),
        cost: String(row[5] || "$0.00"), onHand: parseInt(row[6], 10) || 0, reservedQty: parseInt(row[7], 10) || 0, availableQty: parseInt(row[8], 10) || 0,
        onRevMed: "FALSE", revMedPrice: String(row[10] || ""), syncedShopify: String(row[13] || "FALSE"), 
        category: String(row[14] || ""), status: String(row[15] || "INACTIVE"), parentRef: String(row[16] || ""), uomMult: parseInt(row[17], 10) || 1,
        shelf: String(row[18] || ""), shopifyCategory: String(row[19] || "Medical Supplies")
      });
    }
  }

  var repoOwner = "ASP-Seyfors"; 
  var repoName = GITHUB_REPO_NAME; // <--- SAFE!
  var filePath = "database.json";
  var token = PropertiesService.getScriptProperties().getProperty("GITHUB_PAT");
  
  if (!token) return;

  var getUrl = "https://api.github.com/repos/" + repoOwner + "/" + repoName + "/contents/" + filePath;
  var headers = { "Authorization": "token " + token, "Accept": "application/vnd.github.v3+json" };
  
  try {
    var getResponse = UrlFetchApp.fetch(getUrl, { method: "GET", headers: headers, muteHttpExceptions: true });
    var sha = "";
    if (getResponse.getResponseCode() === 200) { sha = JSON.parse(getResponse.getContentText()).sha; }

    var jsonString = JSON.stringify(db, null, 2);
    var base64Content = Utilities.base64Encode(Utilities.newBlob(jsonString).getBytes());
    var payload = { message: "Automated Daily Database Export", content: base64Content, sha: sha };

    UrlFetchApp.fetch(getUrl, { method: "PUT", headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true });
  } catch (err) {}
}

/*
function patchLegacySessions() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive");
  let data = sheet.getDataRange().getValues();
  
  let patchedCount = 0;

  // Start at i=1 to skip the header. 
  for (let i = 1; i < data.length; i++) {
    let payloadStr = data[i][6]; // Column G (JSON)
    let workflow = (data[i][4] || '').toUpperCase(); // Column E (Status)
    
    // SKIP THE STOCKTAKE ROW (It is truncated and will crash the parser)
    if (String(data[i][3]).includes("Stocktake")) {
      Logger.log("Skipping Stocktake Row " + (i+1) + " due to known size limit.");
      continue;
    }

    if (payloadStr && workflow === 'COMPLETED') {
      try {
        let session = JSON.parse(payloadStr);
        let updated = false;
        let wType = (session.workflowType || '').toUpperCase();
        let sName = (session.sessionName || '').toUpperCase();
        
        session.scannedObjects.forEach(item => {
          let act = item.actionTag || '';
          
          // 1. Add missing sessionId
          if (!item.sessionId) {
            item.sessionId = session.id;
            updated = true;
          }
          
          // 2. Add missing orderNum for Pack/Reserve
          if ((act === 'Reserved' || act === 'Pack & Ship') && !item.orderNum) {
            item.orderNum = session.orderNum || '';
            updated = true;
          }
          
          // 3. Fix blank customerTags during Packing/Reserving
          if (!item.customerTag && (wType.includes('PACK') || wType.includes('RESERVE') || act === 'Pack & Ship' || act === 'Reserved')) {
            let baseCust = sName.split('(')[0].split('-')[0].trim();
            if (baseCust && !baseCust.includes('HISTORICAL')) {
              item.customerTag = baseCust;
              updated = true;
            }
          }
        });
        
        // Save back to the spreadsheet if changes were made
        if (updated) {
          sheet.getRange(i + 1, 7).setValue(JSON.stringify(session));
          patchedCount++;
        }
      } catch (err) {
        Logger.log(`Skipping Row ${i + 1} due to JSON error: ${err.message}`);
      }
    }
  }
  
  Logger.log(`Successfully patched ${patchedCount} sessions.`);
}

function restoreMissingStocktakeItems() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName(SHEET_AUDIT);
  let archiveSheet = ss.getSheetByName("Archive");
  
  if (!auditSheet || !archiveSheet) {
    Logger.log("Missing required sheets.");
    return;
  }
  
  let auditData = auditSheet.getDataRange().getDisplayValues();
  let missingItems = [];
  let stocktakeCount = 0;
  
  // Find all rows belonging to the original stocktake
  for (let i = 1; i < auditData.length; i++) {
    if (auditData[i][2] === "Warehouse Stocktake (FULL-INV)") {
      stocktakeCount++;
      
      // Skip the first 117 items (which are safely in Part 1)
      if (stocktakeCount > 117) {
        let expVal = auditData[i][6] || "NO_EXP";
        if (expVal.indexOf(' ') > -1) expVal = expVal.split(' ')[0]; // Clean timestamps
        
        missingItems.push({
          actionTag: auditData[i][8] || "Inventory",
          gtin: "",
          ref: String(auditData[i][4]).toUpperCase().trim(),
          lot: String(auditData[i][5]).toUpperCase().trim() || "NO_LOT",
          exp: expVal,
          mfr: "N/A",
          desc: "Recovered from Audit Log",
          price: "$0.00",
          qty: parseInt(auditData[i][7], 10) || 1,
          rawScanLines: [],
          isNew: false,
          customerTag: "",
          itemNote: "",
          sessionId: "1786665600979", // New ID for Part 2
          orderNum: ""
        });
      }
    }
  }
  
  if (missingItems.length === 0) {
    Logger.log("No missing items found. You may have already run this script.");
    return;
  }
  
  // Build the Part 2 Payload
  let part2Payload = {
    id: "1786665600979",
    status: "Completed",
    userName: "Thomas",
    sessionName: "Warehouse Stocktake (FULL-INV) Part 2",
    orderNum: "",
    workflowType: "Selective Stocktake", // Selective prevents it from wiping Part 1
    dateStr: "2026.08.14",
    startStr: "09:52 PM",
    manifestEnabled: false,
    expectedManifest: [],
    scannedObjects: missingItems,
    pendingNewItems: [],
    pendingUpdates: [],
    lastUpdated: 1786665600979
  };
  
  // Insert into Cloud Archive
  archiveSheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss"), 
    "'1786665600979", 
    "Thomas", 
    "Warehouse Stocktake (FULL-INV) Part 2", 
    "Completed", 
    "2026.08.14", 
    JSON.stringify(part2Payload)
  ]);
  
  SpreadsheetApp.flush();
  Logger.log(`Success! Recovered ${missingItems.length} missing items and appended Part 2 to the Cloud Archive.`);
}*/