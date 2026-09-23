// ============================================================================================================================
// ============================================================================================================================
//
//                                 ENVIRONMENT SPECIFIC DETAILS DO NOT OVERWRITE THIS SECTION
//
// ============================================================================================================================
// ============================================================================================================================
//                                            SYSTEM USERS (DISTRIBUTION LIST)
// ============================================================================================================================
const APP_USERS = [
  "thomas@alliedsurgicalproducts.com"
  // "jessica@alliedsurgicalproducts.com",
  // "trey@alliedsurgicalproducts.com",
  // "alecia@alliedsurgicalproducts.com"
];
const ADMIN_EMAIL = "thomas@alliedsurgicalproducts.com";
// ============================================================================================================================
//                                                 GITHUB CONFIGURATIONS
// ============================================================================================================================
const GITHUB_REPO_NAME = "ASP-IMS-DEMO";
const GITHUB_DEFAULT_PRODUCT_IMAGE_URL = "https://asp-seyfors.github.io/ASP-IMS-DEMO/ASP_Box_Web_RGB_DEMO.png";
// ============================================================================================================================
//                                              FEDEX SANDBOX CONFIGURATIONS
// ============================================================================================================================
const FEDEX_CONFIG = {
  API_KEY: "l7672edb8ec5ee4867a1869d438115a7a0",
  SECRET_KEY: "8dfc1d2dcc01441f8f95191ecdb120d9",
  ACCOUNT_NUMBER: "740980114",
  BASE_URL: "https://apis-sandbox.fedex.com",
  
  // Shipping Origin Details
  SHIPPER_COMPANY: "Allied Surgical Products",
  SHIPPER_PERSON: "ASP Shipping Dept",
  SHIPPER_PHONE: "7273303360",
  SHIPPER_STREET: "4914 Flora Ave",
  SHIPPER_CITY: "Holiday",
  SHIPPER_STATE: "FL",
  SHIPPER_ZIP: "34690",
  
  // Printer & Tracking Defaults
  LABEL_STOCK_TYPE: "STOCK_4X6", // Use "PAPER_85X11_TOP_HALF_LABEL" for standard desktop printers
  ENV_NOTE: "Demo Label generated." // Change to "Live Label generated." in Prod
};
// ============================================================================================================================
//                                             UPS SANDBOX CONFIGURATIONS
// ============================================================================================================================
const UPS_CONFIG = {
  CLIENT_ID: "1afxK2TF7TT99PBQ6kc9AUSQVtoz544nAvkmAAnPY6EuqARK",
  CLIENT_SECRET: "HGtrnzysQoCo93aw46ivF696XN91efQbH0hYkN3qArBLI23gzgoF5JSm6ThS5oDK",
  ACCOUNT_NUMBER: "E73F37",
  BASE_URL: "https://wwwcie.ups.com" // UPS Sandbox Environment
};
// ============================================================================================================================
//                                             QUICKBOOKS CONFIGURATIONS
// ============================================================================================================================
var QBO_CONFIG = {
  CLIENT_ID: "ABpFa5hwg9TF6QodG0Fp0ncQPzPZ9odblkcWIOEeWZpAL3eDjY",
  CLIENT_SECRET: "LiEcoIfXJxcFqBAjxKnMD411KkD1Vj5wzP6jtbmn",
  REFRESH_TOKEN: "RT1-12-H0-1798769474pryh6chd8wd9l59bvmqd",
  REALM_ID: "9341457908287228", // QBO Demo Sandbox ID
  ENVIRONMENT: "sandbox" // "sandbox" or "production"
};
// ============================================================================================================================
//                                       ORDERS FEED SHEET NAME GLOBAL VARIABLE
// ============================================================================================================================
const FEED_SHEET_NAME = "ASP_IMS_DEMO_FEED"; // ASP_IMS_TEST_FEED, ASP_IMS_DEMO_FEED, or ASP_IMS_ORDERS_FEED
// ============================================================================================================================
//                                             SHOPIFY GLOBAL VARIABLES
// ============================================================================================================================
const SHOPIFY_DEFAULT_CATEGORY = "Medical Supplies";
const SHOPIFY_API_VERSION = "2024-10";
const SHOPIFY_DEFAULT_TAXONOMY_GID = "gid://shopify/TaxonomyCategory/bi-19-7";
const SHOPIFY_BUG_TAG = "Shopify Demo API Trace";
const SHOPIFY_BUG_ENV = "Demo Sandbox";

// ============================================================================================================================
// ============================================================================================================================
//
//                  NON-ENVIRONMENT SPECIFIC GLOBAL VARIABLES (FROM HERE BELOW CAN BE COPIED FOR DEPLOYMENT)
//
// ============================================================================================================================
// ==========================================
// AUTOMATED ON-HAND STOCK REPORT MESSAGE
// ==========================================
const COMPANY_NAME = "Allied Surgical Products";
var CUSTOM_EMAIL_MESSAGE = "Please reach out to your Sales Representative or visit our website at <a href='https://alliedsurgicalproducts.com' style='color: #0277bd; text-decoration: none; font-weight: bold;'>alliedsurgicalproducts.com</a>.";

// ==========================================
// APPS SCRIPT GOLBAL VARIABLES
// ==========================================
// ✨ NEW: Shipping & Tracking Tabs
const SHEET_MFRS = "Manufacturers";
const SHEET_ITEMS = "Items";
const SHEET_BUNDLES = "Bundles";
const SHEET_SUPPLIERS = "Suppliers";
const SHEET_INCOMING = "Incoming";
const SHEET_DMG = "Damaged";
const SHEET_ALLOCATIONS = "Allocations";
const SHEET_CUSTOMERS = "Customers";
const SHEET_SUBSCRIBERS = "Subscriptions";
const SHEET_SHIPPING_RULES = "Shipping_Info";
const SHEET_OUTGOING = "Outgoing";
const SHEET_ARCHIVE = "Archive";
const SHEET_AUDIT = "Audit_Log";
const SHEET_BUG_REPORTS = "Sync_Log";

const ETHICON_URL = "https://www.ethicon.com/na/epc/code/";

// ==========================================
// DATABASE COLUMN MAPPING (0-INDEXED)
// ==========================================
const DB_MAP = {
  REF: 0, 
  MFR: 1, 
  DESC: 2, 
  GTIN: 3, 
  PRICE: 4, 
  COST: 5,
  ON_HAND: 6, 
  RESERVED: 7, 
  AVAILABLE: 8, 
  ON_REVMED: 9, 
  REVMED_PRICE: 10,
  ON_DOTMED: 11, 
  DOTMED_PRICE: 12, 
  SHOPIFY_SYNC: 13, 
  CATEGORY: 14,
  STATUS: 15, 
  PARENT_REF: 16, 
  UOM_MULT: 17, 
  SHELF: 18, 
  SHOPIFY_CAT: 19,
  WEIGHT: 20, 
  DIM_L: 21, 
  DIM_W: 22, 
  DIM_H: 23
};

function doGet(e) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;

    if (action === "CHECK_VERSION") {
      let file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
      return ContentService.createTextOutput(JSON.stringify({ status: "success", lastUpdated: file.getLastUpdated().getTime() })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "VERIFY_USER") {
        let sheet = ss.getSheetByName("Users");
        if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Users tab missing."})).setMimeType(ContentService.MimeType.JSON);
        
        let data = sheet.getDataRange().getValues();
        let emailToFind = (e.parameter.email || "").toLowerCase().trim();
        let userProfile = { role: "GUEST", name: "Guest", disabled: false };
        
        // Loop through starting at row 2 (skipping headers)
        for (let i = 1; i < data.length; i++) {
            let rowEmail = String(data[i][1]).toLowerCase().trim();
            if (rowEmail === emailToFind) {
                userProfile.name = String(data[i][0]);
                userProfile.role = String(data[i][2]).toUpperCase().trim();
                
                // ✨ NEW: Read Theme and Font (Defaults applied if blank)
                userProfile.theme = String(data[i][5] || 'twilight').toLowerCase().trim();
                userProfile.font = String(data[i][6] || 'large').toLowerCase().trim();
                
                if (data[i][4]) {
                    userProfile.disabled = true;
                }
                break;
            }
        }
        return ContentService.createTextOutput(JSON.stringify({status: "success", profile: userProfile})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: ETHICON WEBSITE AUTO-SCRAPER ---
    if (action === "FETCH_ETHICON") {
      let ref = (e.parameter.ref || "").trim().toLowerCase();
      if (!ref) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "No REF provided."})).setMimeType(ContentService.MimeType.JSON);
      
      let url = ETHICON_URL + encodeURIComponent(ref) + "?lang=en-default";
      
      try {
         let res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
         if (res.getResponseCode() >= 400) throw new Error("Product not found on Ethicon.");
         
         let html = res.getContentText();
         let regex = /Description\s*<\/(?:td|th|div|span)>[\s\S]*?<(?:td|div|span)[^>]*>([\s\S]*?)<\/(?:td|div|span)>/i;
         let descMatch = html.match(regex);
         
         if (descMatch && descMatch[1]) {
             let cleanDesc = descMatch[1].replace(/(<([^>]+)>)/gi, "").replace(/&nbsp;/gi, " ").trim();
             
             // ✨ FIX: Convert HTML Entities (like &quot;) into normal text symbols
             cleanDesc = cleanDesc.replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'").replace(/&reg;/gi, '®').replace(/&trade;/gi, '™');
             
             return ContentService.createTextOutput(JSON.stringify({status: "success", desc: cleanDesc})).setMimeType(ContentService.MimeType.JSON);
         } else {
             throw new Error("Could not locate the Description table on the webpage.");
         }
      } catch (err) {
         return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.message})).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- NEW: QBO FEED HANDLER ---
    if (action === "GET_QBO_FEED") {
      let qboSheet = ss.getSheetByName("QBO_Feed");
      let payload = { stagedSessions: {} };
      
      if (qboSheet) {
        let data = qboSheet.getDataRange().getValues();
        let currentSessionKey = "";
        let lastCustomer = "SHELF";
        let lastPO = "NA";
        
        for (let i = 1; i < data.length; i++) {
          let row = data[i];
          
          let rawSession = String(row[1] || '').trim(); // Shifted for Timestamp in 0
          let rawCustomer = String(row[2] || '').trim();
          let rawPO = String(row[3] || '').trim();
          let sku = String(row[4] || '').trim().toUpperCase();
          let qty = parseInt(row[5], 10) || 1;
          let statusVal = String(row[6] || '').trim().toUpperCase(); 

          if (rawSession !== "") {
            currentSessionKey = rawSession.match(/^\d+$/) ? "Session #" + rawSession : rawSession;
            lastCustomer = "SHELF";
            lastPO = "NA";
          }
          if (rawCustomer !== "") lastCustomer = rawCustomer;
          if (rawPO !== "") lastPO = rawPO;

          if (currentSessionKey && sku) {
            if (!payload.stagedSessions[currentSessionKey]) {
              payload.stagedSessions[currentSessionKey] = {
                items: [],
                isCompleted: (statusVal === "COMPLETED")
              };
            }
            
            let cTag = (lastCustomer && lastCustomer !== 'SHELF') ? lastCustomer + (lastPO !== 'NA' && lastPO !== '' ? ' - ' + lastPO : '') : '';
            payload.stagedSessions[currentSessionKey].items.push({ customerTag: cTag, sku: sku, qty: qty });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "GET_SESSION") {
      let targetId = e.parameter.id;
      let archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);
      if (archiveSheet && targetId) {
        let data = archiveSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          let rowId = String(data[i][1]).replace(/'/g, '').trim();
          if (rowId === targetId && data[i][6]) {
            return ContentService.createTextOutput(data[i][6]).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Session payload not found."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "GET_SUBSCRIBERS") {
        let sheet = ss.getSheetByName(SHEET_SUBSCRIBERS);
        if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: "success", subs: []})).setMimeType(ContentService.MimeType.JSON);
        let sData = sheet.getDataRange().getValues();
        let subs = [];
        for (let i = 1; i < sData.length; i++) {
           // Target sData[i][4] for the 5th column
           if(sData[i][1]) subs.push({ name: sData[i][0], email: sData[i][1], freq: sData[i][2], status: sData[i][3], categories: sData[i][4] || "" });
        }
        return ContentService.createTextOutput(JSON.stringify({status: "success", subs: subs})).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "GET_ALLOCATIONS") {
      let allocations = [];
      let allocSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
      if (allocSheet && allocSheet.getLastRow() > 1) {
        let rows = allocSheet.getRange(2, 1, allocSheet.getLastRow() - 1, allocSheet.getLastColumn()).getValues();
        rows.forEach(r => {
          if (r[0] && r[2]) { allocations.push({ customerName: r[0], orderNum: r[1], ref: r[2], lot: r[3], exp: r[4], qty: r[5], sessionId: r[6] }); }
        });
      }
      let dmgSheet = ss.getSheetByName(SHEET_DMG);
      if (dmgSheet && dmgSheet.getLastRow() > 1) {
        let rows = dmgSheet.getRange(2, 1, dmgSheet.getLastRow() - 1, dmgSheet.getLastColumn()).getValues();
        rows.forEach(r => {
          if (r[0] && r[2]) { allocations.push({ customerName: r[0], orderNum: r[1], ref: r[2], lot: r[3], exp: r[4], qty: r[5], sessionId: r[6] }); }
        });
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", allocations: allocations })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'GET_AUDIT_LOG') {
      let logSheet = ss.getSheetByName(SHEET_AUDIT);
      let data = logSheet.getDataRange().getDisplayValues(); 
      let headers = data[0];
      let logs = [];
      for (let i = 1; i < data.length; i++) {
        let rowObj = {};
        for (let j = 0; j < headers.length; j++) { rowObj[headers[j]] = data[i][j]; }
        logs.push(rowObj);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: logs })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: DEEP ARCHIVE LOT TRACEABILITY ---
    if (action === "SEARCH_DEEP_ARCHIVE") {
      let targetLot = (e.parameter.lot || "").toUpperCase().trim();
      let archiveId = PropertiesService.getScriptProperties().getProperty("ARCHIVE_SHEET_ID");
      if (!archiveId) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Archive ID not configured in Script Properties."})).setMimeType(ContentService.MimeType.JSON);
      
      let coldSs = SpreadsheetApp.openById(archiveId);
      let sheets = coldSs.getSheets();
      let foundLogs = [];
      
      for (let i = 0; i < sheets.length; i++) {
        let sheetName = sheets[i].getName();
        if (sheetName.indexOf(SHEET_AUDIT) !== -1) {
          let data = sheets[i].getDataRange().getDisplayValues();
          if (data.length <= 1) continue;
          let headers = data[0];
          for (let j = 1; j < data.length; j++) {
            let rLot = String(data[j][5] || "").toUpperCase(); // Lot is column F
            if (rLot && rLot.indexOf(targetLot) !== -1) {
              let rowObj = {};
              for (let k = 0; k < headers.length; k++) { rowObj[headers[k]] = data[j][k]; }
              foundLogs.push(rowObj);
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: foundLogs })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: FETCH PENDING SHIPMENTS FOR TRACEABILITY ---
    if (action === "GET_PENDING_SHIPMENTS") {
      let outSheet = ss.getSheetByName(SHEET_OUTGOING);
      let inSheet = ss.getSheetByName(SHEET_INCOMING);
      let pendingOut = [], pendingIn = [];
      
      if (outSheet) {
        let outData = outSheet.getDataRange().getValues();
        for (let i = 1; i < outData.length; i++) {
          if (String(outData[i][6]).toUpperCase() === "PENDING") { // Col G is Status
            pendingOut.push({ date: outData[i][0], partner: outData[i][1], po: outData[i][2], carrier: outData[i][3], tracking: outData[i][4] });
          }
        }
      }
      if (inSheet) {
        let inData = inSheet.getDataRange().getValues();
        for (let i = 1; i < inData.length; i++) {
          if (String(inData[i][6]).toUpperCase() !== "DELIVERED" && String(inData[i][6]) !== "") { // Col G is Status
            pendingIn.push({ date: inData[i][0], partner: inData[i][1], po: inData[i][2], carrier: inData[i][3], tracking: inData[i][4] });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", outgoing: pendingOut, incoming: pendingIn })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- CHUNKED DATABASE DOWNLOAD ENGINE (OPTIMIZED MATRIX) ---
    if (action === "SYNC_DATABASE_CHUNKED") {
      let page = parseInt(e.parameter.page || 1, 10);
      let pageSize = 1500; // Increased batch size due to compression
      let items = [];
      
      let itemsSheet = ss.getSheetByName(SHEET_ITEMS);
      if (itemsSheet) {
        let itemsData = itemsSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < itemsData.length; i++) {
          let row = itemsData[i];
          if (row[DB_MAP.REF]) items.push([
            String(row[DB_MAP.REF]), 
            String(row[DB_MAP.MFR]), 
            String(row[DB_MAP.DESC]), 
            String(row[DB_MAP.GTIN]), 
            String(row[DB_MAP.PRICE]),
            String(row[DB_MAP.COST] || "$0.00"), 
            parseInt(row[DB_MAP.ON_HAND], 10) || 0, 
            parseInt(row[DB_MAP.RESERVED], 10) || 0,    
            parseInt(row[DB_MAP.AVAILABLE], 10) || 0, 
            String(row[DB_MAP.ON_REVMED] || "FALSE"), 
            String(row[DB_MAP.REVMED_PRICE] || ""),        
            String(row[DB_MAP.ON_DOTMED] || "FALSE"), 
            String(row[DB_MAP.DOTMED_PRICE] || ""),
            String(row[DB_MAP.SHOPIFY_SYNC] || "FALSE"), 
            String(row[DB_MAP.CATEGORY] || ""), 
            String(row[DB_MAP.STATUS] || "ACTIVE"),
            String(row[DB_MAP.PARENT_REF] || ""), 
            parseInt(row[DB_MAP.UOM_MULT], 10) || 1, 
            String(row[DB_MAP.SHELF] || ""),
            String(row[DB_MAP.SHOPIFY_CAT] || SHOPIFY_DEFAULT_CATEGORY),
            String(row[DB_MAP.WEIGHT] || "0"), 
            String(row[DB_MAP.DIM_L] || "0"), 
            String(row[DB_MAP.DIM_W] || "0"), 
            String(row[DB_MAP.DIM_H] || "0")
          ]);
        }
      }

      let uomSheet = ss.getSheetByName(SHEET_BUNDLES);
      if (uomSheet) {
        let uomData = uomSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < uomData.length; i++) {
          let row = uomData[i];
          if (row[DB_MAP.REF]) items.push([
            String(row[DB_MAP.REF]), 
            String(row[DB_MAP.MFR]), 
            String(row[DB_MAP.DESC]), 
            String(row[DB_MAP.GTIN]), 
            String(row[DB_MAP.PRICE]),
            String(row[DB_MAP.COST] || "$0.00"), 0, 0, 0, "FALSE", "", "FALSE", "", 
            String(row[DB_MAP.SHOPIFY_SYNC] || "FALSE"), 
            String(row[DB_MAP.CATEGORY] || ""), 
            String(row[DB_MAP.STATUS] || "ACTIVE"),
            String(row[DB_MAP.PARENT_REF] || ""), 
            parseInt(row[DB_MAP.UOM_MULT], 10) || 1, 
            String(row[DB_MAP.SHELF] || ""),
            String(row[DB_MAP.SHOPIFY_CAT] || SHOPIFY_DEFAULT_CATEGORY),
            String(row[DB_MAP.WEIGHT] || "0"), 
            String(row[DB_MAP.DIM_L] || "0"), 
            String(row[DB_MAP.DIM_W] || "0"), 
            String(row[DB_MAP.DIM_H] || "0")
          ]);
        }
      }
      
      let totalPages = Math.ceil(items.length / pageSize);
      let chunkedItems = items.slice((page - 1) * pageSize, page * pageSize);
      let dbPayload = { matrix: chunkedItems }; // Use 'matrix' key
      
      if (page === 1) {
         dbPayload.customers = []; dbPayload.suppliers = []; dbPayload.vendors = [];
         dbPayload.customerAliases = {}; dbPayload.supplierAliases = {};
         dbPayload.shippingRules = {}; // ✨ NEW
         
         let custSheet = ss.getSheetByName(SHEET_CUSTOMERS);
         if (custSheet) { 
           let custData = custSheet.getDataRange().getValues(); 
           for (let i = 1; i < custData.length; i++) { 
             let primary = String(custData[i][0]).trim();
             if (primary) { 
               dbPayload.customers.push(primary); 
               // ✨ NEW: Read the rest of the row for aliases
               for (let j = 1; j < custData[i].length; j++) {
                 let alias = String(custData[i][j]).trim().toUpperCase();
                 if (alias) dbPayload.customerAliases[alias] = primary;
               }
             } 
           } 
         }
         
         let supSheet = ss.getSheetByName(SHEET_SUPPLIERS);
         if (supSheet) { 
           let supData = supSheet.getDataRange().getValues(); 
           for (let i = 1; i < supData.length; i++) { 
             let primary = String(supData[i][0]).trim();
             if (primary) { 
               dbPayload.suppliers.push(primary); 
               // ✨ NEW: Read the rest of the row for aliases
               for (let j = 1; j < supData[i].length; j++) {
                 let alias = String(supData[i][j]).trim().toUpperCase();
                 if (alias) dbPayload.supplierAliases[alias] = primary;
               }
             } 
           } 
         }
         
         let mfrSheet = ss.getSheetByName(SHEET_MFRS);
         if (mfrSheet) { let mfrData = mfrSheet.getDataRange().getValues(); for (let i = 1; i < mfrData.length; i++) { if (mfrData[i][0]) dbPayload.vendors.push(String(mfrData[i][0]).trim()); } }

         // ✨ NEW: Fetch Customer Shipping Info (With Alias Cross-Referencing)
         let shipSheet = ss.getSheetByName(SHEET_SHIPPING_RULES);
         if (shipSheet) {
             let shipData = shipSheet.getDataRange().getValues();
             for (let i = 1; i < shipData.length; i++) {
                 let rawCust = String(shipData[i][1]).trim().toUpperCase(); // Col B: Contact Name (System ID)
                 if (rawCust) {
                     // Check if this name is an alias. If so, link it to the Primary Name.
                     let primaryCust = dbPayload.customerAliases[rawCust] || rawCust;

                     dbPayload.shippingRules[primaryCust] = {
                         formalCompany: String(shipData[i][0]).trim(), // Exact name for FedEx label
                         contactName: String(shipData[i][1]).trim(),   // Col B: Contact Name
                         email: String(shipData[i][2]).trim(),         // Col C: Contact Email
                         phone: String(shipData[i][3]).trim(),         // Col D: Contact Phone
                         address: String(shipData[i][4]).trim(),       // Col E: Shipping Address
                         method: String(shipData[i][5]).trim(),        // Col F: Preferred Shipping Method
                         account: String(shipData[i][6]).trim(),       // Col G: Shipping Account Number
                         notes: String(shipData[i][7]).trim()          // Col H: Notes
                     };
                 }
             }
         }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
         status: "success", page: page, totalPages: totalPages, db: dbPayload
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- DEFAULT ACTION: SYNC_DIRECTORY ---
    let db = { items: [], customers: [], suppliers: [], vendors: [] };
    let sessionsLite = [];
    
    let archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);
    if (archiveSheet) {
      let archiveData = archiveSheet.getDataRange().getValues();
      for (let i = 1; i < archiveData.length; i++) {
        let rowId = String(archiveData[i][1]).replace(/'/g, '').trim();
        if (rowId) {
          sessionsLite.push({ id: rowId, userName: String(archiveData[i][2]), sessionName: String(archiveData[i][3]), status: String(archiveData[i][4]), dateStr: String(archiveData[i][5]), isCloud: true });
        }
      }
    }
    
    let custSheet = ss.getSheetByName(SHEET_CUSTOMERS);
    if (custSheet) { let custData = custSheet.getDataRange().getValues(); for (let i = 1; i < custData.length; i++) { if (custData[i][0]) db.customers.push(String(custData[i][0])); } }
    
    let supSheet = ss.getSheetByName(SHEET_SUPPLIERS);
    if (supSheet) { let supData = supSheet.getDataRange().getValues(); for (let i = 1; i < supData.length; i++) { if (supData[i][0]) db.suppliers.push(String(supData[i][0])); } }
    
    let mfrSheet = ss.getSheetByName(SHEET_MFRS);
    if (mfrSheet) { let mfrData = mfrSheet.getDataRange().getValues(); for (let i = 1; i < mfrData.length; i++) { if (mfrData[i][0]) db.vendors.push(String(mfrData[i][0])); } }

    if (action === "GET_BUG_REPORTS") {
      let sheet = ss.getSheetByName(SHEET_BUG_REPORTS);
      let reports = [];
      if (sheet && sheet.getLastRow() > 1) {
        let data = sheet.getDataRange().getDisplayValues();
        for (let i = data.length - 1; i > 0; i--) { 
          reports.push({ timestamp: data[i][0], user: data[i][1], version: data[i][2], session: data[i][3], workflow: data[i][4], desc: data[i][5], env: data[i][6] });
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", reports: reports })).setMimeType(ContentService.MimeType.JSON);
    }      
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", archive: sessionsLite, db: db })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  let lock = LockService.getScriptLock();
  lock.tryLock(30000); 

  try {
    let raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
    if (!raw) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No postData received" })).setMimeType(ContentService.MimeType.JSON);

    let data = JSON.parse(raw);
    let ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- NEW: QBO POST HANDLERS ---
    if (data.action === "FETCH_QBO") {
      let result = fetchOpenInvoicesFromQBO();
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === "QBO_WRITEBACK") {
      var sessionObj = data.payload || {};
      
      // 1. Stamp the QBO_Feed tab with "COMPLETED"
      var liveOrderNum = sessionObj.orderNum || "";
      var feedSheet = ss.getSheetByName("QBO_Feed");
      
      if (feedSheet && liveOrderNum) {
        var feedData = feedSheet.getDataRange().getValues();
        var currentSession = "";
        
        for (var i = 1; i < feedData.length; i++) {
          var rawSession = String(feedData[i][1] || '').trim(); // Adjusted index to 1 for Session Name
          if (rawSession !== "") {
            currentSession = rawSession;
          }
          if (currentSession.toUpperCase().includes(liveOrderNum.toUpperCase())) {
            feedSheet.getRange(i + 1, 7).setValue("COMPLETED"); // Adjusted index to 7 for Status
          }
        }
        SpreadsheetApp.flush();
      }

      // 2. Execute the QBO Invoice update
      var writeBackResult = writeBackToQboInvoice(sessionObj);
      
      return ContentService.createTextOutput(JSON.stringify(writeBackResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: FEDEX SHIPMENT ENDPOINT ---
    if (data.action === "CREATE_SHIPMENT") {
      let result = createFedExShipment(data.payload || {});
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: MANUAL TRACKING LOGGER ---
    if (data.action === "LOG_MANUAL_TRACKING") {
      let outSheet = ss.getSheetByName(SHEET_OUTGOING);
      if (outSheet) {
        let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
        let p = data.payload;
        
        // ✨ NEW: Generate Hyperlink based on Carrier
        let trackLink = "";
        let safeTrack = String(p.trackingNumber).trim();
        if (String(p.carrier).toUpperCase().indexOf("UPS") > -1) {
            trackLink = '=HYPERLINK("https://www.ups.com/track?track=yes&trackNums=' + safeTrack + '", "Track")';
        } else if (String(p.carrier).toUpperCase().indexOf("FEDEX") > -1) {
            trackLink = '=HYPERLINK("https://www.fedex.com/fedextrack/?trknbr=' + safeTrack + '", "Track")';
        }

        let dims = p.dimL ? (p.dimL + "x" + p.dimW + "x" + p.dimH) : "";
        let weightStr = p.totalWeight ? p.totalWeight + " lbs" : "";

        outSheet.appendRow([
          timestamp,            // A: Date
          p.customerName,       // B: Customer Name
          p.orderNum,           // C: Invoice / PO
          p.carrier,            // D: Carrier
          p.trackingNumber,     // E: Tracking Number
          "Yes",                // F: Tracking Email
          "Pending",            // G: Status
          "",                   // H: ETA
          "",                   // I: Final Invoice Sent?
          "Client provided label.", // J: Notes
          trackLink,            // K: Carrier Link
          "",                   // L: ETA This Week?
          "",                   // M: Delay Reason
          weightStr,            // N: Weight
          dims                  // O: Dimensions
        ]);
        SpreadsheetApp.flush();
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: UPS SHIPMENT ENDPOINT STUB ---
    if (data.action === "CREATE_UPS_SHIPMENT") {
      // Will be wired up fully once the sandbox token fully propagates
      return ContentService.createTextOutput(JSON.stringify({
        status: "error", 
        message: "UPS Sandbox credentials are still propagating. Please use manual portal."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ✨ NEW: SUBSCRIBER ENDPOINT
    if (data.action === "UPDATE_SUBSCRIBER") {
        let sheet = ss.getSheetByName(SHEET_SUBSCRIBERS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_SUBSCRIBERS);
            sheet.appendRow(["Customer Name", "Email Address", "Frequency", "Status"]);
        }
        let sData = sheet.getDataRange().getValues();
        let p = data.payload;
        let found = false;
        
        for (let i = 1; i < sData.length; i++) {
            if (String(sData[i][1]).trim().toLowerCase() === String(p.email).trim().toLowerCase()) {
                // Write 5 values starting at column 1
                sheet.getRange(i+1, 1, 1, 5).setValues([[p.name, p.email, p.freq, p.status, p.categories || ""]]);
                found = true;
                break;
            }
        }
        if (!found) {
            // Append exactly 5 values
            sheet.appendRow([p.name, p.email, p.freq, p.status, p.categories || ""]);
            
            // ✨ FIX: Fire welcome email to new active subscribers
            if (String(p.status).toUpperCase() === 'ACTIVE') {
                let welcomeHtml = "<div style='font-family:Arial,sans-serif; color:#333; max-width:600px; margin:auto; border:1px solid #e0e0e0; padding:20px;'>" +
                  "<div style='text-align:center; margin-bottom: 20px; border-bottom: 3px solid #0277bd; padding-bottom: 15px;'>" +
                  "<h2 style='color:#0277bd; margin: 5px 0;'>Allied Surgical Products</h2>" +
                  "</div>" +
                  "<p>Hello " + p.name + ",</p>" +
                  "<p>You have been successfully added to our Automated Stock Report distribution list. You will receive our latest inventory updates on a <strong>" + p.freq + "</strong> basis.</p>" +
                  "<p style='color:#555; font-size:12px;'><em>Please add this email address to your safe senders list or mark it as 'Not Spam' to ensure our reports reach your inbox.</em></p>" +
                  "<p>If you would like to change your frequency or be removed from this list, please contact us at <a href='mailto:sales@alliedsurgicalproducts.com'>sales@alliedsurgicalproducts.com</a>.</p>" +
                  "</div>";

                MailApp.sendEmail({
                    to: p.email,
                    subject: "Welcome to the ASP Stock Report List",
                    htmlBody: welcomeHtml,
                    name: "Allied Surgical Products"
                });
            }
        }
        
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "ARCHIVE_SESSION") {
      let sheet = ss.getSheetByName(SHEET_ARCHIVE);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_ARCHIVE);
        sheet.appendRow(["Timestamp", "Session ID", "User Name", "Session Name", "Status", "Date", "JSON Payload"]);
      }
      
      let sess = data.payload || {};
      let baseSessionName = sess.sessionName || "Unnamed Session";
      let dateParts = (sess.dateStr || '').toString().trim().split(/[\.\-\/]/);
      let colADate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
      let colFDate = sess.dateStr || '';

      if (dateParts.length === 3) {
        let y, m, d;
        if (dateParts[0].length === 4) { y = parseInt(dateParts[0], 10); m = parseInt(dateParts[1], 10) - 1; d = parseInt(dateParts[2], 10); } 
        else { m = parseInt(dateParts[0], 10) - 1; d = parseInt(dateParts[1], 10); y = parseInt(dateParts[2], 10); }
        colFDate = y + "." + Utilities.formatString("%02d", m + 1) + "." + Utilities.formatString("%02d", d);
      }

      // ✨ NEW: AUTO-SPLIT MASSIVE PAYLOADS (Protects against 50k char cell limit)
      let MAX_ITEMS_PER_ROW = 100;
      let scannedItems = sess.scannedObjects || [];
      let totalChunks = Math.ceil(scannedItems.length / MAX_ITEMS_PER_ROW) || 1;
      
      let existingIdsClean = sheet.getRange("B:B").getValues().flat().map(v => String(v).replace(/'/g, '').trim());
      
      for (let chunkIdx = 0; chunkIdx < totalChunks; chunkIdx++) {
        let isChunked = totalChunks > 1;
        let chunkItems = scannedItems.slice(chunkIdx * MAX_ITEMS_PER_ROW, (chunkIdx + 1) * MAX_ITEMS_PER_ROW);
        
        let sessionName = isChunked ? `${baseSessionName} (Part ${chunkIdx + 1} of ${totalChunks})` : baseSessionName;
        let targetIdClean = String(sess.id || Date.now()).replace(/'/g, '').trim();
        if (isChunked) targetIdClean += `_P${chunkIdx + 1}`; // Create unique ID for the chunk
        
        // Clone and trim the session object specifically for this chunk
        let chunkSess = JSON.parse(JSON.stringify(sess));
        chunkSess.sessionName = sessionName;
        chunkSess.id = targetIdClean;
        chunkSess.scannedObjects = chunkItems;
        
        let rowIndex = existingIdsClean.indexOf(targetIdClean);
        let isNewArchiveRow = (rowIndex === -1); 
        
        let rowData = [ colADate, "'" + targetIdClean, sess.userName || "User", sessionName, sess.status || "Completed", colFDate, JSON.stringify(chunkSess) ];

        if (!isNewArchiveRow) sheet.getRange(rowIndex + 1, 1, 1, 7).setValues([rowData]);
        else sheet.appendRow(rowData);

        // Process Audit Log for this specific chunk
        if (isNewArchiveRow && sess.status === "Completed" && chunkItems.length > 0) {
          let auditSheetName = SHEET_AUDIT;
          let auditSheet = ss.getSheetByName(auditSheetName);
          if (!auditSheet) {
            auditSheet = ss.insertSheet(auditSheetName);
            auditSheet.appendRow(["Timestamp", "User", "Session / Reason", "Workflow", "REF / SKU", "Lot", "Exp Date", "Qty Moved", "Destination / Action"]);
            auditSheet.getRange("A1:I1").setFontWeight("bold");
          }
          let auditRows = [];
          chunkItems.forEach(item => {
            let cLot = (item.lot === 'N/A' || item.lot === 'NA' || item.lot === 'NO_LOT') ? '' : item.lot;
            let cExp = (item.exp === 'N/A' || item.exp === 'NA' || item.exp === 'NO_EXP') ? '' : item.exp;
            auditRows.push([colADate, sess.userName || "User", sessionName, sess.workflowType || "General", item.ref || "UNKNOWN", cLot, cExp, item.qty || 0, item.actionTag || "Inventory"]);
          });
          if (auditRows.length > 0) auditSheet.getRange(auditSheet.getLastRow() + 1, 1, auditRows.length, 9).setValues(auditRows);
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "SYNC_ALLOCATIONS") {
      let standardSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
      let dmgSheet = ss.getSheetByName(SHEET_DMG);
      
      if (!standardSheet) { standardSheet = ss.insertSheet(SHEET_ALLOCATIONS); standardSheet.appendRow(["Customer Name", "Order Number", "REF / SKU", "Lot", "Exp Date", "Reserved Qty", "Session ID"]); standardSheet.getRange("A1:G1").setFontWeight("bold"); }
      if (!dmgSheet) { dmgSheet = ss.insertSheet(SHEET_DMG); dmgSheet.appendRow(["Customer Name", "Note", "REF / SKU", "Lot", "Exp Date", "Reserved Qty", "Session ID"]); dmgSheet.getRange("A1:G1").setFontWeight("bold"); }
      
      if (standardSheet.getLastRow() > 1) standardSheet.getRange(2, 1, standardSheet.getLastRow() - 1, standardSheet.getLastColumn()).clearContent();
      if (dmgSheet.getLastRow() > 1) dmgSheet.getRange(2, 1, dmgSheet.getLastRow() - 1, dmgSheet.getLastColumn()).clearContent();
      
      let standardRows = [];
      let dmgRows = [];
      let allocs = data.allocations || {};
      let aliases = data.aliases || {}; // ✨ FIX: Pull the exact proper case from the frontend payload
      
      if (!Array.isArray(allocs)) {
        for (let cust in allocs) {
          let resolvedCust = aliases[cust] || cust; 
          
          // ✨ FIX: Do the check using toUpperCase, but push the preserved 'resolvedCust' to the sheet!
          let isDamaged = (String(resolvedCust).toUpperCase().trim() === "ASP DAMAGED INVENTORY");
          let targetArray = isDamaged ? dmgRows : standardRows;

          for (let ref in allocs[cust]) {
            let itemData = allocs[cust][ref];
            
            if (itemData && itemData.details && itemData.details.length > 0) {
              itemData.details.forEach(det => {
                if (det.qty > 0) {
                  // Uses resolvedCust exactly as it was mapped by the alias engine
                  targetArray.push([resolvedCust, det.orderNum || "", ref, det.lot || "", det.exp || "", det.qty, det.sessionId || ""]);
                }
              });
            } 
            else if (itemData && itemData.qty > 0) {
              targetArray.push([resolvedCust, "", ref, "", "", itemData.qty, ""]);
            }
            else if (typeof itemData === 'number' && itemData > 0) {
              targetArray.push([resolvedCust, "", ref, "", "", itemData, ""]);
            }
          }
        }
      }

      if (standardRows.length > 0) standardSheet.getRange(2, 1, standardRows.length, 7).setValues(standardRows);
      if (dmgRows.length > 0) dmgSheet.getRange(2, 1, dmgRows.length, 7).setValues(dmgRows);
      
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: SAVE ADDRESS BOOK ENTRY ---
    if (data.action === "SAVE_SHIPPING_INFO") {
      let sheet = ss.getSheetByName(SHEET_SHIPPING_RULES);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_SHIPPING_RULES);
        sheet.appendRow(["Company Name", "Contact Name", "Contact Email", "Contact Phone", "Shipping Address", "Preferred Shipping Method", "Shipping Account Number", "Shipping/Payment Terms Notes"]);
      }
      let p = data.payload;
      let sData = sheet.getDataRange().getValues();
      let found = false;
      let targetSystemName = String(p.customerName).toUpperCase().trim(); // Short name from the Session

      for (let i = 1; i < sData.length; i++) {
        let rowContactName = String(sData[i][1]).toUpperCase().trim(); // Column B
        
        // Match strictly on the Contact Name (Column B)
        if (rowContactName === targetSystemName) {
          // Update existing row. We force p.customerName into the Contact Name slot to lock the relationship
          sheet.getRange(i+1, 1, 1, 8).setValues([[p.formalCompany, p.customerName, p.email, p.phone, p.address, p.method, p.account, p.notes]]);
          found = true;
          break;
        }
      }
      if (!found) {
        // Append new row
        sheet.appendRow([p.formalCompany, p.customerName, p.email, p.phone, p.address, p.method, p.account, p.notes]);
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // =======================================================================
    // --- NEW: API INTEGRATION ENDPOINTS (SHOPIFY, FEDEX, UPS) ---
    // =======================================================================
    // --- NEW: SAVE USER PREFERENCES ---
    if (data.action === "SAVE_USER_PREFS") {
      let sheet = ss.getSheetByName("Users");
      if (sheet) {
        let sData = sheet.getDataRange().getValues();
        let targetEmail = String(data.payload.email).toLowerCase().trim();
        for (let i = 1; i < sData.length; i++) {
          if (String(sData[i][1]).toLowerCase().trim() === targetEmail) {
             sheet.getRange(i + 1, 6).setValue(data.payload.theme); // Column F
             sheet.getRange(i + 1, 7).setValue(data.payload.font);  // Column G
             break;
          }
        }
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 1. SHOPIFY: LIVE SEED TEST
    if (data.action === "TEST_SHOPIFY_CONNECTION") {
      let result = testShopifyConnection(data.payload || []);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. SHOPIFY: LIVE WAREHOUSE SYNC
    if (data.action === "SYNC_SHOPIFY_SANDBOX") {
      let result = syncShopifySandbox(data.payload || []);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === "SYNC_LOCAL_DB") {
      let custSheet = ss.getSheetByName(SHEET_CUSTOMERS);
      if (custSheet) {
        let existingCusts = custSheet.getDataRange().getValues().flat().map(c => String(c).toUpperCase());
        (data.payload.customers || []).forEach(c => {
          if (c && !existingCusts.includes(c.toUpperCase())) { custSheet.appendRow([c]); existingCusts.push(c.toUpperCase()); }
        });
      }
      
      let supSheet = ss.getSheetByName(SHEET_SUPPLIERS);
      if (supSheet) {
        let existingSups = supSheet.getDataRange().getValues().flat().map(s => String(s).toUpperCase());
        (data.payload.suppliers || []).forEach(s => {
          if (s && !existingSups.includes(s.toUpperCase())) { supSheet.appendRow([s]); existingSups.push(s.toUpperCase()); }
        });
      }

      let mfrSheet = ss.getSheetByName(SHEET_MFRS);
      if (mfrSheet) {
        let existingMfrs = mfrSheet.getDataRange().getValues().flat().map(m => String(m).toUpperCase());
        (data.payload.vendors || []).forEach(m => {
          if (m && !existingMfrs.includes(m.toUpperCase())) { mfrSheet.appendRow([m]); existingMfrs.push(m.toUpperCase()); }
        });
      }

      let itemSheet = ss.getSheetByName(SHEET_ITEMS);
      if (itemSheet) {
        let itemVals = itemSheet.getDataRange().getValues();
        
        let maxCols = 21; 
      
        // Loop 1: Find the true max columns
        for (let i = 0; i < itemVals.length; i++) {
          if (itemVals[i].length > maxCols) maxCols = itemVals[i].length;
        }
      
        let refMap = {}; 
      
        // Loop 2: Pad rows, inject formulas, AND build the refMap simultaneously
        for (let i = 0; i < itemVals.length; i++) {
          while (itemVals[i].length < maxCols) itemVals[i].push(""); 
          if (i > 0) {
              itemVals[i][DB_MAP.AVAILABLE] = "=G" + (i + 1) + "-H" + (i + 1);
              if (itemVals[i][DB_MAP.REF]) refMap[String(itemVals[i][DB_MAP.REF]).toUpperCase().replace(/'/g, '').trim()] = i; 
          }
        }

        let newRows = [];
        
        (data.payload.items || []).forEach(item => {
          let ref = String(item.ref || item.sku || "").toUpperCase().trim();
          if (!ref) return;
          
          // ROUTE BUNDLES TO UOM TAB
          if (item.parentRef && item.uomMult > 1 && item.parentRef.toUpperCase() !== ref) {
            let uomSheet = ss.getSheetByName(SHEET_BUNDLES);
            if (uomSheet) {
              let uomVals = uomSheet.getDataRange().getValues();
              let uomMap = {}; 
              for (let i = 1; i < uomVals.length; i++) {
                if (uomVals[i][0]) uomMap[String(uomVals[i][0]).toUpperCase().replace(/'/g, '').trim()] = i;
              }
              if (uomMap[ref] === undefined) {
                uomSheet.appendRow([
                    "'" + ref, item.mfr || "", item.desc || "", item.gtin ? "'" + item.gtin : "", item.price || "$0.00", item.cost || "$0.00",
                    0, 0, 0, "FALSE", "", "", "", 
                    item.syncedShopify || "FALSE", 
                    item.category || "General", 
                    item.status || "ACTIVE", item.parentRef, item.uomMult, item.shelf || "",
                    item.shopifyCategory || SHOPIFY_DEFAULT_CATEGORY,
                    item.weight || "0", item.dimL || "0", item.dimW || "0", item.dimH || "0" // ✨ NEW
                ]);
              }
            }
            return; 
          }
          
          if (refMap[ref] !== undefined) 
          {
            let rowIdx = refMap[ref]; 
            if (item.mfr !== undefined) itemVals[rowIdx][DB_MAP.MFR] = item.mfr;
            if (item.desc !== undefined) itemVals[rowIdx][DB_MAP.DESC] = item.desc;
            if (item.gtin !== undefined) itemVals[rowIdx][DB_MAP.GTIN] = "'" + item.gtin;
            if (item.price && item.price !== "$0.00") itemVals[rowIdx][DB_MAP.PRICE] = item.price;
            if (item.cost && item.cost !== "$0.00") itemVals[rowIdx][DB_MAP.COST] = item.cost;
            if (item.onHand !== undefined) itemVals[rowIdx][DB_MAP.ON_HAND] = item.onHand; 
            if (item.reservedQty !== undefined) itemVals[rowIdx][DB_MAP.RESERVED] = item.reservedQty; 
            if (item.onRevMed !== undefined) itemVals[rowIdx][DB_MAP.ON_REVMED] = item.onRevMed;
            if (item.revMedPrice !== undefined) itemVals[rowIdx][DB_MAP.REVMED_PRICE] = item.revMedPrice;
            if (item.onDotMed !== undefined) itemVals[rowIdx][DB_MAP.ON_DOTMED] = item.onDotMed;
            if (item.dotMedPrice !== undefined) itemVals[rowIdx][DB_MAP.DOTMED_PRICE] = item.dotMedPrice;
            if (item.syncedShopify !== undefined) itemVals[rowIdx][DB_MAP.SHOPIFY_SYNC] = item.syncedShopify;
            if (item.category !== undefined) itemVals[rowIdx][DB_MAP.CATEGORY] = item.category;
            if (item.status !== undefined) itemVals[rowIdx][DB_MAP.STATUS] = item.status;
            if (item.parentRef !== undefined && item.parentRef.toUpperCase() !== ref) itemVals[rowIdx][DB_MAP.PARENT_REF] = item.parentRef; 
            if (item.uomMult !== undefined) itemVals[rowIdx][DB_MAP.UOM_MULT] = item.uomMult; 
            if (item.shelf !== undefined) itemVals[rowIdx][DB_MAP.SHELF] = item.shelf;
            if (item.shopifyCategory !== undefined) itemVals[rowIdx][DB_MAP.SHOPIFY_CAT] = item.shopifyCategory;
            if (item.weight !== undefined) itemVals[rowIdx][DB_MAP.WEIGHT] = item.weight;
            if (item.dimL !== undefined) itemVals[rowIdx][DB_MAP.DIM_L] = item.dimL;
            if (item.dimW !== undefined) itemVals[rowIdx][DB_MAP.DIM_W] = item.dimW;
            if (item.dimH !== undefined) itemVals[rowIdx][DB_MAP.DIM_H] = item.dimH;
          } else {
            let nextRowNumber = itemVals.length + newRows.length + 1;
            let newRow = new Array(maxCols).fill("");
            
            newRow[DB_MAP.REF] = "'" + ref;
            newRow[DB_MAP.MFR] = item.mfr || "";
            newRow[DB_MAP.DESC] = item.desc || "";
            newRow[DB_MAP.GTIN] = item.gtin ? "'" + item.gtin : "";
            newRow[DB_MAP.PRICE] = item.price || "$0.00";
            newRow[DB_MAP.COST] = item.cost || "$0.00";
            newRow[DB_MAP.ON_HAND] = item.onHand || 0;
            newRow[DB_MAP.RESERVED] = item.reservedQty || 0;
            newRow[DB_MAP.AVAILABLE] = "=G" + nextRowNumber + "-H" + nextRowNumber;
            newRow[DB_MAP.ON_REVMED] = item.onRevMed || "FALSE";
            newRow[DB_MAP.REVMED_PRICE] = item.revMedPrice || "";
            newRow[DB_MAP.ON_DOTMED] = item.onDotMed || "FALSE";
            newRow[DB_MAP.DOTMED_PRICE] = item.dotMedPrice || "";
            newRow[DB_MAP.SHOPIFY_SYNC] = item.syncedShopify || "FALSE";
            newRow[DB_MAP.CATEGORY] = item.category || "General";
            newRow[DB_MAP.STATUS] = item.status || "ACTIVE";
            newRow[DB_MAP.PARENT_REF] = (item.parentRef && item.parentRef.toUpperCase() !== ref) ? item.parentRef : "";
            newRow[DB_MAP.UOM_MULT] = item.uomMult || 1;
            newRow[DB_MAP.SHELF] = item.shelf || "";
            newRow[DB_MAP.SHOPIFY_CAT] = item.shopifyCategory || SHOPIFY_DEFAULT_CATEGORY;
            newRow[DB_MAP.WEIGHT] = item.weight || "0";
            newRow[DB_MAP.DIM_L] = item.dimL || "0";
            newRow[DB_MAP.DIM_W] = item.dimW || "0";
            newRow[DB_MAP.DIM_H] = item.dimH || "0";
            
            newRows.push(newRow);
          }
        });
        
        if (itemVals.length > 0) itemSheet.getRange(1, 1, itemVals.length, maxCols).setValues(itemVals);
        if (newRows.length > 0) itemSheet.getRange(itemVals.length + 1, 1, newRows.length, maxCols).setValues(newRows);
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: PRICE ALERT ENDPOINT ---
    if (data.action === "PRICE_ALERT") {
      let payload = data.payload || {};
      let emailBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto;'>" +
        "<h2 style='color:#f57f17; border-bottom:2px solid #f57f17; padding-bottom:8px; margin-bottom:15px;'>💰 Item Price Update Alert</h2>" +
        "<p style='margin: 4px 0;'><strong>User:</strong> " + (payload.user || "Unknown") + "</p>" +
        "<p style='margin: 4px 0;'><strong>REF / SKU:</strong> " + (payload.ref || "Unknown") + "</p>" +
        "<p style='margin: 4px 0;'><strong>Previous Price:</strong> <span style='color:#c62828; text-decoration:line-through;'>" + (payload.oldPrice || "$0.00") + "</span></p>" +
        "<p style='margin: 4px 0;'><strong>New Price:</strong> <span style='color:#2e7d32; font-weight:bold; font-size:16px;'>" + (payload.newPrice || "$0.00") + "</span></p>" +
        "<div style='background:#f9f9f9; border:1px solid #ccc; border-left:4px solid #f57f17; padding:12px; margin-top:15px; font-size:13px; color:#555;'>" +
        "This change was submitted via the Database Editor and will permanently reflect in the master catalog on the next cloud upload." +
        "</div></div>";

      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: "💰 ASP Price Update Alert - " + (payload.ref || "Unknown"),
        htmlBody: emailBody,
        name: "ASP System Alerts"
      });
      
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- NEW: BUG REPORT ENDPOINT ---
    if (data.action === "REPORT_BUG") {
      let sheet = ss.getSheetByName(SHEET_BUG_REPORTS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_BUG_REPORTS);
        sheet.appendRow(["Timestamp", "User", "App Version", "Active Session", "Active Workflow", "Description", "Environment"]);
        sheet.getRange("A1:G1").setFontWeight("bold");
      }
      let payload = data.payload || {};
      let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
      
      sheet.appendRow([
        timestamp, 
        payload.userName || "Unknown", 
        payload.appVersion || "Unknown", 
        payload.sessionName || "None", 
        payload.workflowType || "None", 
        payload.description || "",
        payload.environment || "Production"
      ]);
      
      // --- NEW: AUTOMATED EMAIL ALERT ---
      let emailBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto;'>" +
        "<h2 style='color:#c62828; border-bottom:2px solid #c62828; padding-bottom:8px; margin-bottom:15px;'>🐞 New System Bug Report</h2>" +
        "<p style='margin: 4px 0;'><strong>User:</strong> " + (payload.userName || "Unknown") + "</p>" +
        "<p style='margin: 4px 0;'><strong>Environment:</strong> " + (payload.environment || "Production") + " (" + (payload.appVersion || "Unknown") + ")</p>" +
        "<p style='margin: 4px 0;'><strong>Session:</strong> " + (payload.sessionName || "None") + "</p>" +
        "<p style='margin: 4px 0;'><strong>Workflow:</strong> " + (payload.workflowType || "None") + "</p>" +
        "<div style='background:#f9f9f9; border:1px solid #ccc; border-left:4px solid #c62828; padding:12px; margin-top:15px; font-family:monospace; font-size:13px; color:#c62828; white-space:pre-wrap;'>" + (payload.description || "") + "</div>" +
        "</div>";

      MailApp.sendEmail({
        to: ADMIN_EMAIL, // Change this if you want it sent elsewhere
        subject: "🚨 ASP IMS Bug Report (" + (payload.environment || "Production") + ")",
        htmlBody: emailBody,
        name: "ASP System Alerts"
      });
      // ----------------------------------
      
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: SEND DEPLOYMENT UPDATE EMAIL ---
    if (data.action === "SEND_UPDATE_EMAIL") {
      let result = sendUpdateEmail(data.payload || {});
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // --- NEW: APPS SCRIPT AUTO-EXPORTER ---
    if (data.action === "EXPORT_APPS_SCRIPT") {
      let result = exportAppsScriptToDrive();
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function sendUpdateEmail(data) {
  try {
    let oldVer = data.oldVersion || "Previous";
    let newVer = data.newVersion || "Latest";
    let notes = data.notes || "General UI improvements and bug fixes.";

    let subject = `🚀 ASP IMS Update: ${newVer}`;
    
    let htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        <div style="background-color: #7b1fa2; color: #ffffff; padding: 20px; text-align: center;">
          <h2 style="margin: 0; font-size: 1.5rem;">ASP Inventory Management System Update</h2>
          <p style="margin: 5px 0 0 0; font-size: 1.1rem; color: #e1bee7;">Version ${oldVer} ➔ <strong>Version ${newVer}</strong></p>
        </div>
        
        <div style="padding: 20px; color: #333; line-height: 1.6;">
          <p style="font-size: 1.05rem;">Hello Team,</p>
          <p style="font-size: 1.05rem;">A new version of the ASP Inventory Scanner has just been deployed to Production. Please review the release notes and follow the steps below to ensure your devices are running the latest code.</p>
          
          <div style="background-color: #f3e5f5; padding: 15px; border-left: 4px solid #7b1fa2; border-radius: 4px; margin: 20px 0;">
            <h4 style="margin: 0 0 10px 0; color: #4a148c;">Release Notes:</h4>
            <p style="margin: 0; white-space: pre-wrap; font-size: 0.95rem;">${notes}</p>
          </div>

          <h3 style="color: #7b1fa2; border-bottom: 2px solid #f3e5f5; padding-bottom: 5px; margin-top: 30px;">🖥️ Desktop / Laptop Instructions</h3>
          <ul style="padding-left: 20px; font-size: 0.95rem;">
            <li style="margin-bottom: 8px;"><strong>Installed App:</strong> If you installed the app to your computer, uninstall it first, then go to the URL and click the install icon in the address bar again.</li>
            <li style="margin-bottom: 8px;"><strong>Windows PC:</strong> While on the scanner page, press <strong>Ctrl + F5</strong> (or Ctrl + Shift + R) to perform a hard refresh.</li>
            <li style="margin-bottom: 8px;"><strong>Mac / Apple:</strong> While on the scanner page, press <strong>Cmd + Shift + R</strong>.</li>
          </ul>

          <h3 style="color: #7b1fa2; border-bottom: 2px solid #f3e5f5; padding-bottom: 5px; margin-top: 30px;">📱 Mobile Instructions</h3>
          <ul style="padding-left: 20px; font-size: 0.95rem;">
            <li style="margin-bottom: 8px;"><strong>iPhone (iOS):</strong> Delete the bookmark from your home screen. Open Safari, go to the scanner URL, clear your Safari cache if necessary, tap the Share icon, and select "Add to Home Screen".</li>
            <li style="margin-bottom: 8px;"><strong>Android:</strong> Remove the app from your home screen. Open Chrome, go to the scanner URL, clear the site data if necessary, tap the 3-dot menu, and select "Install App" or "Add to Home Screen".</li>
          </ul>
          
          <p style="margin-top: 30px; font-size: 1.05rem;">If you experience any issues, please submit a Bug Report through the app's User Menu.</p>
          <p style="font-size: 1.05rem; margin-bottom: 0;">Thank you,<br><strong>Thomas</strong></p>
        </div>
      </div>
    `;

    // Send the email using MailApp to avoid scope authorization errors
    MailApp.sendEmail({
      to: APP_USERS[0],
      bcc: APP_USERS.join(","),
      subject: subject,
      htmlBody: htmlBody,
      name: "ASP System Admin"
    });

    return { status: "success" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function exportAppsScriptToDrive() {
  try {
    let scriptId = ScriptApp.getScriptId();
    
    // ✨ THE FIX: Route the extraction through the Google Drive Export API 
    // instead of the restricted Apps Script Management API.
    let mime = encodeURIComponent("application/vnd.google-apps.script+json");
    let url = "https://www.googleapis.com/drive/v3/files/" + scriptId + "/export?mimeType=" + mime;
    
    // Because this script uses DriveApp below, this token automatically 
    // holds the required Google Drive security permissions.
    let token = ScriptApp.getOAuthToken();
    
    let options = {
      "method": "get",
      "headers": { "Authorization": "Bearer " + token },
      "muteHttpExceptions": true
    };
    
    let res = UrlFetchApp.fetch(url, options);
    
    if (res.getResponseCode() >= 400) {
       Logger.log("Google Drive API Error: " + res.getContentText());
       return;
    }
    
    let scriptContent = JSON.parse(res.getContentText());
    if (!scriptContent.files || scriptContent.files.length === 0) {
      Logger.log("No script files found in this project.");
      return;
    }

    let folderId = PropertiesService.getScriptProperties().getProperty("BACKUP_FOLDER_ID");
    if (!folderId) {
       Logger.log("ERROR: BACKUP_FOLDER_ID not set in Script Properties.");
       return;
    }

    let destFolder = DriveApp.getFolderById(folderId);
    let timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
    let backupFolder = destFolder.createFolder("Code_Backup_" + timeStamp);

    let savedFiles = 0;
    scriptContent.files.forEach(function(file) {
      let extension = ".gs";
      if (file.type === "HTML") extension = ".html";
      if (file.type === "JSON") extension = ".json"; // Catches the appsscript.json file
      
      backupFolder.createFile(file.name + extension, file.source);
      savedFiles++;
    });

    Logger.log("SUCCESS! Backed up " + savedFiles + " files to Google Drive Folder.");
  } catch (err) {
    Logger.log("FATAL ERROR: " + err.toString());
  }
}