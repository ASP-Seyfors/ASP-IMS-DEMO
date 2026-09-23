function sendAutomatedEndOfWeekReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEET_AUDIT);
  var allocSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
  var itemsSheet = ss.getSheetByName("Items");
  var archiveSheet = ss.getSheetByName("Archive");
  
  if (!auditSheet) return;

  // 1. Calculate dates (Last 7 days)
  var today = new Date();
  var cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 7);
  cutoffDate.setHours(0,0,0,0);
  
  // 2. Build Price Map from Items Sheet to calculate Revenue
  var priceMap = {};
  if (itemsSheet) {
    var itemsData = itemsSheet.getDataRange().getValues();
    for (var i = 1; i < itemsData.length; i++) {
      var ref = String(itemsData[i][DB_MAP.REF]).trim();
      var priceStr = String(itemsData[i][DB_MAP.PRICE] || "").replace(/[^0-9.-]+/g, "");
      var price = parseFloat(priceStr) || 0;
      if (ref) priceMap[ref] = price;
    }
  }

  // 3. Read Archive JSON Payloads for New REFs
  var newlyAddedRefs = {};
  if (archiveSheet) {
    var archData = archiveSheet.getDataRange().getValues();
    for (var i = 1; i < archData.length; i++) {
      var rowDate = new Date(archData[i][0]); // Timestamp in column A
      if (rowDate < cutoffDate) continue;
      
      var sessionName = String(archData[i][3] || "").toUpperCase();
      var isTest = ["TEST SUPPLIER", "ASP_INTERNAL", "ASP_TESTER", "ASP_TESTER2", "TESTER"].some(function(kw) { return sessionName.indexOf(kw) !== -1; }) || sessionName === "TEST" || sessionName.indexOf("TEST ") === 0;
      if (isTest) continue;

      var payloadStr = archData[i][6]; // JSON Payload
      if (payloadStr) {
        try {
          var payload = JSON.parse(payloadStr);
          if (payload.pendingNewItems && payload.pendingNewItems.length > 0) {
            payload.pendingNewItems.forEach(function(item) {
              if (item.ref) newlyAddedRefs[item.ref] = true;
            });
          }
        } catch(e) {}
      }
    }
  }
  
  // 4. Read Audit Log for KPIs and Revenue
  var data = auditSheet.getDataRange().getValues();
  var headers = data.shift();
  
  var dateIdx = headers.indexOf("Timestamp");
  var sessionIdx = headers.indexOf("Session / Reason");
  var workflowIdx = headers.indexOf("Workflow");
  var qtyIdx = headers.indexOf("Qty Moved");
  var refIdx = headers.indexOf("REF / SKU");
  
  var sessionsSet = {};
  var outboundSessions = {};
  var totalItems = 0;
  var uniqueRefs = {};
  var totalRevenue = 0;
  
  var testKeywords = ["TEST SUPPLIER", "ASP_INTERNAL", "ASP_TESTER", "ASP_TESTER2", "TESTER"];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowDate = new Date(row[dateIdx]);
    if (rowDate < cutoffDate) continue;
    
    var sessionName = String(row[sessionIdx] || "");
    var upperSession = sessionName.toUpperCase();
    
    // Exclude tests
    var isTest = testKeywords.some(function(kw) { return upperSession.indexOf(kw) !== -1; }) || upperSession === "TEST" || upperSession.indexOf("TEST ") === 0;
    if (isTest) continue;
    
    var workflow = String(row[workflowIdx] || "");
    var qty = Math.abs(parseInt(row[qtyIdx]) || 0);
    var ref = String(row[refIdx] || "");
    
    sessionsSet[sessionName] = true;
    if (workflow.indexOf('Packing') !== -1 || workflow.indexOf('Pack & Ship') !== -1) {
      outboundSessions[sessionName] = true;
      totalRevenue += (qty * (priceMap[ref] || 0)); // Mathematically calculate using live Item price
    }
    if (ref && ref !== 'N/A') {
      totalItems += qty;
      uniqueRefs[ref] = true;
    }
  }
  
  // 5. Read Active Allocations for Customer Bins
  var allocHtml = "";
  if (allocSheet) {
    var allocData = allocSheet.getDataRange().getValues();
    var aHeaders = allocData.shift();
    var cIdx = aHeaders.indexOf("Customer Name"); 
    var aRefIdx = aHeaders.indexOf("REF / SKU");
    var aQtyIdx = aHeaders.indexOf("Reserved Qty");
    
    if (cIdx !== -1 && aRefIdx !== -1 && aQtyIdx !== -1) {
      var bins = {};
      for (var j = 0; j < allocData.length; j++) {
        var c = allocData[j][cIdx];
        var r = allocData[j][aRefIdx];
        var q = parseInt(allocData[j][aQtyIdx]) || 0;
        if (c && r && q > 0) {
          if (!bins[c]) bins[c] = { total: 0, items: [] };
          bins[c].items.push({ref: r, qty: q});
          bins[c].total += q;
        }
      }
      
      var customers = Object.keys(bins).sort();
      if (customers.length === 0) {
         allocHtml = "<p style='color:#777; font-size:12px; font-style:italic;'>No active reservations found.</p>";
      } else {
         for (var k = 0; k < customers.length; k++) {
           var cust = customers[k];
           var bin = bins[cust];
           var rowsHtml = "";
           for (var m = 0; m < bin.items.length; m++) {
             rowsHtml += "<tr><td style='padding:4px 8px; border:1px solid #eee;'>" + bin.items[m].ref + "</td><td style='padding:4px 8px; border:1px solid #eee; text-align:center; font-weight:bold;'>" + bin.items[m].qty + "</td></tr>";
           }
           allocHtml += "<div style='margin-bottom: 15px; border: 1px solid #bfe0fb; border-radius: 4px; padding: 10px; background: #fff;'><div style='color: #0277bd; font-weight: bold; font-size: 14px; margin-bottom: 6px; border-bottom: 1px solid #eee; padding-bottom: 4px;'>" + cust + " <span style='float:right; color:#2e7d32;'>Total Items: " + bin.total + "</span></div><table style='width:100%; border-collapse: collapse; font-size: 12px;'><tr style='background:#f0f8ff;'><th style='text-align:left; padding:4px 8px;'>REF</th><th style='text-align:center; padding:4px 8px;'>Reserved Qty</th></tr>" + rowsHtml + "</table></div>";
         }
      }
    }
  }
  
  var newRefsArr = Object.keys(newlyAddedRefs).sort();
  var newRefsHtml = newRefsArr.length > 0 ? "<ul style='margin-top:0;'><li>" + newRefsArr.join("</li><li>") + "</li></ul>" : "<p style='color:#777; font-size:12px; font-style:italic;'>No new REFs added this week.</p>";
  
  // Build Email HTML
  var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto;'>";
  htmlBody += "<h2 style='color:#0277bd; border-bottom:2px solid #0277bd; padding-bottom:10px; margin-bottom:4px;'>ASP Weekly Warehouse Report</h2>";
  htmlBody += "<p style='margin-top:0; font-size:12px; color:#555;'>Generated: <strong>" + today.toLocaleDateString() + "</strong></p>";
  htmlBody += "<table style='width:100%; text-align:center; margin-bottom:20px; border-collapse:collapse;'><tr>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Total Items Scanned</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + totalItems + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Unique REFs</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(uniqueRefs).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>New REFs</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + newRefsArr.length + "</span></td>";
  htmlBody += "</tr><tr>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Sessions Logged</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(sessionsSet).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Orders Packed</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(outboundSessions).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Revenue (Packed)</strong><br><span style='font-size:18px; color:#2e7d32; font-weight:bold;'>$" + totalRevenue.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + "</span></td>";
  htmlBody += "</tr></table>";
  
  htmlBody += "<h3 style='color:#0277bd; margin-bottom:6px;'>✨ NEW REFS ADDED THIS WEEK</h3>" + newRefsHtml;
  htmlBody += "<h3 style='color:#0277bd; margin-bottom:6px;'>📦 CURRENT CUSTOMER RESERVED BINS</h3>" + allocHtml;
  htmlBody += "</div>";
  
  // Send Email (Update this string with the actual email alias or comma-separated addresses)
  var emailTo = ADMIN_EMAIL; 
  
  MailApp.sendEmail({
    to: emailTo,
    subject: "📊 ASP Weekly Inventory Report - " + today.toLocaleDateString(),
    htmlBody: htmlBody,
    name: COMPANY_NAME // Branded!
  });
}

function sendAutomatedEndOfDayReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEET_AUDIT);
  var allocSheet = ss.getSheetByName(SHEET_ALLOCATIONS);
  var itemsSheet = ss.getSheetByName("Items");
  var archiveSheet = ss.getSheetByName("Archive");
  
  if (!auditSheet) return;

  // 1. Calculate dates (Today only)
  var today = new Date();
  var cutoffDate = new Date();
  cutoffDate.setHours(0,0,0,0);
  
  // 2. Build Price Map from Items Sheet to calculate Revenue
  var priceMap = {};
  if (itemsSheet) {
    var itemsData = itemsSheet.getDataRange().getValues();
    for (var i = 1; i < itemsData.length; i++) {
      var ref = String(itemsData[i][DB_MAP.REF]).trim();
      var priceStr = String(itemsData[i][DB_MAP.PRICE] || "").replace(/[^0-9.-]+/g, "");
      var price = parseFloat(priceStr) || 0;
      if (ref) priceMap[ref] = price;
    }
  }

  // 3. Read Archive JSON Payloads for New REFs
  var newlyAddedRefs = {};
  if (archiveSheet) {
    var archData = archiveSheet.getDataRange().getValues();
    for (var i = 1; i < archData.length; i++) {
      var rowDate = new Date(archData[i][0]);
      if (rowDate < cutoffDate) continue;
      
      var sessionName = String(archData[i][3] || "").toUpperCase();
      var isTest = ["TEST SUPPLIER", "ASP_INTERNAL", "ASP_TESTER", "ASP_TESTER2", "TESTER"].some(function(kw) { return sessionName.indexOf(kw) !== -1; }) || sessionName === "TEST" || sessionName.indexOf("TEST ") === 0;
      if (isTest) continue;

      var payloadStr = archData[i][6];
      if (payloadStr) {
        try {
          var payload = JSON.parse(payloadStr);
          if (payload.pendingNewItems && payload.pendingNewItems.length > 0) {
            payload.pendingNewItems.forEach(function(item) {
              if (item.ref) newlyAddedRefs[item.ref] = true;
            });
          }
        } catch(e) {}
      }
    }
  }
  
  // 4. Read Audit Log for KPIs and Revenue
  var data = auditSheet.getDataRange().getValues();
  var headers = data.shift();
  
  var dateIdx = headers.indexOf("Timestamp");
  var sessionIdx = headers.indexOf("Session / Reason");
  var workflowIdx = headers.indexOf("Workflow");
  var qtyIdx = headers.indexOf("Qty Moved");
  var refIdx = headers.indexOf("REF / SKU");
  
  var sessionsSet = {};
  var outboundSessions = {};
  var totalItems = 0;
  var uniqueRefs = {};
  var totalRevenue = 0;
  
  var testKeywords = ["TEST SUPPLIER", "ASP_INTERNAL", "ASP_TESTER", "ASP_TESTER2", "TESTER"];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowDate = new Date(row[dateIdx]);
    if (rowDate < cutoffDate) continue;
    
    var sessionName = String(row[sessionIdx] || "");
    var upperSession = sessionName.toUpperCase();
    
    var isTest = testKeywords.some(function(kw) { return upperSession.indexOf(kw) !== -1; }) || upperSession === "TEST" || upperSession.indexOf("TEST ") === 0;
    if (isTest) continue;
    
    var workflow = String(row[workflowIdx] || "");
    var qty = Math.abs(parseInt(row[qtyIdx]) || 0);
    var ref = String(row[refIdx] || "");
    
    sessionsSet[sessionName] = true;
    if (workflow.indexOf('Packing') !== -1 || workflow.indexOf('Pack & Ship') !== -1) {
      outboundSessions[sessionName] = true;
      totalRevenue += (qty * (priceMap[ref] || 0)); 
    }
    if (ref && ref !== 'N/A') {
      totalItems += qty;
      uniqueRefs[ref] = true;
    }
  }
  
  // 5. Read Active Allocations for Customer Bins
  var allocHtml = "";
  if (allocSheet) {
    var allocData = allocSheet.getDataRange().getValues();
    var aHeaders = allocData.shift();
    var cIdx = aHeaders.indexOf("Customer Name"); 
    var aRefIdx = aHeaders.indexOf("REF / SKU");
    var aQtyIdx = aHeaders.indexOf("Reserved Qty");
    
    if (cIdx !== -1 && aRefIdx !== -1 && aQtyIdx !== -1) {
      var bins = {};
      for (var j = 0; j < allocData.length; j++) {
        var c = allocData[j][cIdx];
        var r = allocData[j][aRefIdx];
        var q = parseInt(allocData[j][aQtyIdx]) || 0;
        if (c && r && q > 0) {
          if (!bins[c]) bins[c] = { total: 0, items: [] };
          bins[c].items.push({ref: r, qty: q});
          bins[c].total += q;
        }
      }
      
      var customers = Object.keys(bins).sort();
      if (customers.length === 0) {
         allocHtml = "<p style='color:#777; font-size:12px; font-style:italic;'>No active reservations found.</p>";
      } else {
         for (var k = 0; k < customers.length; k++) {
           var cust = customers[k];
           var bin = bins[cust];
           var rowsHtml = "";
           for (var m = 0; m < bin.items.length; m++) {
             rowsHtml += "<tr><td style='padding:4px 8px; border:1px solid #eee;'>" + bin.items[m].ref + "</td><td style='padding:4px 8px; border:1px solid #eee; text-align:center; font-weight:bold;'>" + bin.items[m].qty + "</td></tr>";
           }
           allocHtml += "<div style='margin-bottom: 15px; border: 1px solid #bfe0fb; border-radius: 4px; padding: 10px; background: #fff;'><div style='color: #0277bd; font-weight: bold; font-size: 14px; margin-bottom: 6px; border-bottom: 1px solid #eee; padding-bottom: 4px;'>" + cust + " <span style='float:right; color:#2e7d32;'>Total Items: " + bin.total + "</span></div><table style='width:100%; border-collapse: collapse; font-size: 12px;'><tr style='background:#f0f8ff;'><th style='text-align:left; padding:4px 8px;'>REF</th><th style='text-align:center; padding:4px 8px;'>Reserved Qty</th></tr>" + rowsHtml + "</table></div>";
         }
      }
    }
  }
  
  var newRefsArr = Object.keys(newlyAddedRefs).sort();
  var newRefsHtml = newRefsArr.length > 0 ? "<ul style='margin-top:0;'><li>" + newRefsArr.join("</li><li>") + "</li></ul>" : "<p style='color:#777; font-size:12px; font-style:italic;'>No new REFs added today.</p>";
  
  var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: auto;'>";
  htmlBody += "<h2 style='color:#0277bd; border-bottom:2px solid #0277bd; padding-bottom:10px; margin-bottom:4px;'>ASP Daily Warehouse Report</h2>";
  htmlBody += "<p style='margin-top:0; font-size:12px; color:#555;'>Generated: <strong>" + today.toLocaleDateString() + "</strong></p>";
  htmlBody += "<table style='width:100%; text-align:center; margin-bottom:20px; border-collapse:collapse;'><tr>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Total Items Scanned</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + totalItems + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Unique REFs</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(uniqueRefs).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>New REFs</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + newRefsArr.length + "</span></td>";
  htmlBody += "</tr><tr>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Sessions Logged</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(sessionsSet).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Orders Packed</strong><br><span style='font-size:18px; color:#0277bd; font-weight:bold;'>" + Object.keys(outboundSessions).length + "</span></td>";
  htmlBody += "<td style='padding:10px; border:1px solid #ccc; background:#f0f8ff;'><strong>Revenue (Packed)</strong><br><span style='font-size:18px; color:#2e7d32; font-weight:bold;'>$" + totalRevenue.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + "</span></td>";
  htmlBody += "</tr></table>";
  
  htmlBody += "<h3 style='color:#0277bd; margin-bottom:6px;'>✨ NEW REFS ADDED TODAY</h3>" + newRefsHtml;
  htmlBody += "<h3 style='color:#0277bd; margin-bottom:6px;'>📦 CURRENT CUSTOMER RESERVED BINS</h3>" + allocHtml;
  htmlBody += "</div>";
  
  var emailTo = ADMIN_EMAIL; 
  
  MailApp.sendEmail({
    to: emailTo,
    subject: "📊 ASP Daily Inventory Report - " + today.toLocaleDateString(),
    htmlBody: htmlBody,
    name: COMPANY_NAME
  });
}

function sendAutomatedOnHandReport(triggerFrequency) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itemsSheet = ss.getSheetByName("Items");
  var subSheet = ss.getSheetByName(SHEET_SUBSCRIBERS);
  if (!itemsSheet || !subSheet) return;
  
  // 1. Determine who gets this report based on the trigger (Daily vs Weekly)
  var targetFreq = triggerFrequency || "Weekly"; 
  var subData = subSheet.getDataRange().getValues();
  var recipients = [];
  
  for (var k = 1; k < subData.length; k++) {
      var sEmail = String(subData[k][1] || "").trim();
      var sFreq = String(subData[k][2] || "").trim();
      var sStatus = String(subData[k][3] || "").trim().toUpperCase();
      var sCats = String(subData[k][5] || "").trim(); // Ensure this matches your Category column index
      
      if (sEmail && sStatus === "ACTIVE" && sFreq.toUpperCase() === targetFreq.toUpperCase()) {
          recipients.push({ email: sEmail, categories: sCats });
      }
  }
  
  if (recipients.length === 0) return; // No active subscribers for this run

  var data = itemsSheet.getDataRange().getValues();
  var headers = data.shift(); 
  
  var onHandItems = [];
  var totalUnits = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ref = String(row[DB_MAP.REF]).trim();
    var mfr = String(row[DB_MAP.MFR]).trim() || "UNKNOWN";
    var desc = String(row[DB_MAP.DESC]).trim();
    var priceStr = String(row[DB_MAP.PRICE] || "").replace(/[^0-9.-]+/g, "");
    var price = parseFloat(priceStr) || 0;
    
    var totalQty = parseInt(row[DB_MAP.ON_HAND], 10) || 0;
    var resQty = parseInt(row[DB_MAP.RESERVED], 10) || 0;
    var availQty = totalQty - resQty;

    var category = String(row[DB_MAP.CATEGORY] || "").trim();
    
    // Filter out out-of-stock and unpriced items
    if (availQty > 0) {
      onHandItems.push({
        mfr: mfr,
        ref: ref,
        desc: desc,
        price: row[DB_MAP.PRICE],
        category: category
      });
      totalUnits += availQty;
    }
  }

  // INDIVIDUAL EMAIL LOOP
  for (var r = 0; r < recipients.length; r++) {
    var rec = recipients[r];
    
    // Parse the subscriber's category preferences into an array
    var prefArray = rec.categories.split(',').map(function(c){ return c.trim(); }).filter(function(c){ return c !== ""; });

    // ✨ FIX 1: Filter using 'onHandItems' instead of the undefined variable
    var filteredItems = onHandItems.filter(function(item) {
        if (prefArray.length === 0) return true; // Send all if no preference saved
        
        // Handle items that might have multiple categories themselves (e.g. "Medical, Surgical")
        var itemCats = item.category.split(',').map(function(c){ return c.trim(); });
        
        // If the item has ANY category that matches the user's preferences, include it
        return itemCats.some(function(c) { return prefArray.indexOf(c) !== -1; });
    });

    // Skip this subscriber if their preferred categories are completely out of stock
    if (filteredItems.length === 0) continue;

    // Sort their custom list
    filteredItems.sort(function(a, b) {
      return a.mfr.localeCompare(b.mfr) || a.ref.localeCompare(b.ref);
    });

    var today = new Date();
    var dateStr = today.toLocaleDateString();

    var htmlRows = "";
    // ✨ FIX 2: Loop targets 'filteredItems.length' and 'filteredItems[j]'
    for (var j = 0; j < filteredItems.length; j++) {
      var itm = filteredItems[j];
      var formattedPrice = (itm.price && parseFloat(itm.price) > 0) ? "$" + parseFloat(itm.price).toFixed(2) : "CONTACT US";
      htmlRows += "<tr>";
      // Using white-space:nowrap ensures MFR, REF, and PRICE never line-break.
      htmlRows += "<td style='padding:6px; border:1px solid #ddd; white-space:nowrap;'>" + itm.mfr + "</td>";
      htmlRows += "<td style='padding:6px; border:1px solid #ddd; font-weight:bold; color:#0277bd; white-space:nowrap;'>" + itm.ref + "</td>";
      htmlRows += "<td style='padding:6px; border:1px solid #ddd; font-size:11px;'>" + itm.desc + "</td>";
      htmlRows += "<td style='padding:6px; border:1px solid #ddd; text-align:right; white-space:nowrap;'>" + formattedPrice + "</td>";
      htmlRows += "</tr>";
    }

    var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: auto; border: 1px solid #e0e0e0; padding: 20px;'>";
    
    htmlBody += "<div style='text-align:center; margin-bottom: 20px; border-bottom: 3px solid #0277bd; padding-bottom: 15px;'>";
    htmlBody += "<img src='" + GITHUB_DEFAULT_PRODUCT_IMAGE_URL + "' alt='Allied Surgical Products' style='height: 80px; margin-bottom: 10px;' />";
    htmlBody += "<h2 style='color:#0277bd; margin: 5px 0;'>Allied Surgical Products</h2>";
    htmlBody += "<p style='margin: 2px 0; font-size:12px;'>Local: +1 (727) 330-3360 | Toll-Free: +1 (888) 254-8969</p>";
    htmlBody += "<p style='margin: 2px 0; font-size:12px;'><a href='mailto:sales@alliedsurgicalproducts.com'>sales@alliedsurgicalproducts.com</a></p>";
    htmlBody += "</div>";
    
    htmlBody += "<h3 style='color:#333; margin-bottom:5px;'>📦 ASP On-Hand Stock Catalog (" + targetFreq + ")</h3>";
    htmlBody += "<p style='margin-top:0; font-size:12px; color:#555;'>Generated: <strong>" + dateStr + "</strong></p>";
    
    // ✨ FIX 4: Re-added the missing Summary Box
    htmlBody += "<table style='width:100%; text-align:center; margin-bottom:20px; border-collapse:collapse;'><tr>";
    htmlBody += "<td style='padding:15px; border:1px solid #ccc; background:#e3f2fd;'><strong>Total Catalog Items</strong><br><span style='font-size:20px; color:#0277bd; font-weight:bold;'>" + filteredItems.length + "</span></td>";
    htmlBody += "</tr></table>";

    htmlBody += "<table style='width:100%; border-collapse:collapse; font-size:12px; text-align:left;'>";
    htmlBody += "<tr style='background:#f0f0f0;'><th style='padding:8px; border:1px solid #ccc;'>MFR</th><th style='padding:8px; border:1px solid #ccc;'>REF</th><th style='padding:8px; border:1px solid #ccc;'>DESC</th><th style='padding:8px; border:1px solid #ccc; text-align:right;'>PRICE</th></tr>";
    htmlBody += htmlRows;
    htmlBody += "</table>";
    
    // Injects your global message at the bottom of the table
    htmlBody += "<div style='margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #0277bd; font-size: 13px; color: #555;'>";
    htmlBody += CUSTOM_EMAIL_MESSAGE;
    htmlBody += "</div></div>";

    MailApp.sendEmail({
      to: rec.email,
      subject: "📦 ASP " + targetFreq + " Stock Report (" + dateStr + ")",
      htmlBody: htmlBody,
      name: "Allied Surgical Products"
    });
  }
} 

function sendMonthlyVolumeReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEET_AUDIT);
  var itemsSheet = ss.getSheetByName("Items");
  if (!auditSheet) return;

  // 1. Build Pricing & Cost Maps
  var priceMap = {};
  var costMap = {};
  if (itemsSheet) {
    var itemsData = itemsSheet.getDataRange().getValues();
    for (var i = 1; i < itemsData.length; i++) {
      var ref = String(itemsData[i][DB_MAP.REF]).trim();
      priceMap[ref] = parseFloat(String(itemsData[i][DB_MAP.PRICE] || "").replace(/[^0-9.-]+/g, "")) || 0;
      costMap[ref] = parseFloat(String(itemsData[i][DB_MAP.COST] || "").replace(/[^0-9.-]+/g, "")) || 0;
    }
  }

  // 2. Calculate Date Range (Last 30 Days)
  var today = new Date();
  var cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 30);
  cutoffDate.setHours(0,0,0,0);
  var dateStr = cutoffDate.toLocaleDateString() + " - " + today.toLocaleDateString();

  var data = auditSheet.getDataRange().getValues();
  var headers = data.shift();

  var inboundBySupplier = {};
  var outboundByCustomer = {};
  var totalInboundQty = 0, totalInboundCost = 0;
  var totalOutboundQty = 0, totalOutboundRev = 0;
  
  var testKeywords = ["TEST SUPPLIER", "ASP_INTERNAL", "ASP_TESTER", "ASP_TESTER2", "TESTER"];

  // 3. Process Ledger
  for (var j = 0; j < data.length; j++) {
    var rowDate = new Date(data[j][0]);
    if (rowDate < cutoffDate) continue;

    var sessionName = String(data[j][2] || "");
    var upperSession = sessionName.toUpperCase();
    if (testKeywords.some(function(kw) { return upperSession.indexOf(kw) !== -1; }) || upperSession === "TEST" || upperSession.indexOf("TEST ") === 0) continue;

    var workflow = String(data[j][3] || "").toUpperCase();
    var itemRef = String(data[j][4] || "").trim();
    var qty = Math.abs(parseInt(data[j][7]) || 0);
    var partner = sessionName.split('(')[0].trim() || "UNKNOWN";

    if (workflow.indexOf('RECEIVING') !== -1) {
      if (!inboundBySupplier[partner]) inboundBySupplier[partner] = { qty: 0, cost: 0 };
      inboundBySupplier[partner].qty += qty;
      inboundBySupplier[partner].cost += (qty * (costMap[itemRef] || 0));
      totalInboundQty += qty;
      totalInboundCost += (qty * (costMap[itemRef] || 0));
    } else if (workflow.indexOf('PACKING') !== -1 || workflow.indexOf('PACK & SHIP') !== -1) {
      if (!outboundByCustomer[partner]) outboundByCustomer[partner] = { qty: 0, rev: 0 };
      outboundByCustomer[partner].qty += qty;
      outboundByCustomer[partner].rev += (qty * (priceMap[itemRef] || 0));
      totalOutboundQty += qty;
      totalOutboundRev += (qty * (priceMap[itemRef] || 0));
    }
  }

  // 4. Build CSV Attachment
  var csvContent = "TYPE,PARTNER,TOTAL UNITS,FINANCIAL VALUE\r\n";
  var inHtml = "", outHtml = "";

  var inboundKeys = Object.keys(inboundBySupplier).sort();
  for (var k = 0; k < inboundKeys.length; k++) {
    var sup = inboundKeys[k];
    var dataIn = inboundBySupplier[sup];
    csvContent += '"INBOUND","' + sup + '",' + dataIn.qty + ',"$' + dataIn.cost.toFixed(2) + '"\r\n';
    inHtml += "<tr><td style='padding:6px; border:1px solid #ddd;'>" + sup + "</td><td style='padding:6px; border:1px solid #ddd; text-align:center; font-weight:bold;'>" + dataIn.qty + "</td><td style='padding:6px; border:1px solid #ddd; text-align:right; color:#c62828;'>$" + dataIn.cost.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}) + "</td></tr>";
  }

  var outboundKeys = Object.keys(outboundByCustomer).sort();
  for (var m = 0; m < outboundKeys.length; m++) {
    var cust = outboundKeys[m];
    var dataOut = outboundByCustomer[cust];
    csvContent += '"OUTBOUND","' + cust + '",' + dataOut.qty + ',"$' + dataOut.rev.toFixed(2) + '"\r\n';
    outHtml += "<tr><td style='padding:6px; border:1px solid #ddd;'>" + cust + "</td><td style='padding:6px; border:1px solid #ddd; text-align:center; font-weight:bold;'>" + dataOut.qty + "</td><td style='padding:6px; border:1px solid #ddd; text-align:right; color:#2e7d32;'>$" + dataOut.rev.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}) + "</td></tr>";
  }

  var csvBlob = Utilities.newBlob(csvContent, 'text/csv', 'ASP_Volume_Report_' + today.toLocaleDateString().replace(/\//g, '.') + '.csv');

  // 5. Build HTML Email
  var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: auto;'>";
  htmlBody += "<h2 style='color:#0277bd; border-bottom:2px solid #0277bd; padding-bottom:10px; margin-bottom:15px;'>📈 Monthly Volume & Financial Report</h2>";
  htmlBody += "<p style='margin-top:0; font-size:12px; color:#555;'>Period: <strong>" + dateStr + "</strong></p>";

  htmlBody += "<div style='display:flex; gap:15px; margin-bottom:20px;'>";
  
  htmlBody += "<table style='width:100%; border-collapse:collapse;'><tr style='background:#fff3e0;'><th colspan='2' style='padding:10px; border:1px solid #ccc; color:#e65100; font-size:16px;'>📥 Inbound Receiving</th></tr>";
  htmlBody += "<tr><td style='padding:10px; border:1px solid #ccc;'>Total Units Received:</td><td style='padding:10px; border:1px solid #ccc; font-weight:bold; text-align:right;'>" + totalInboundQty + "</td></tr>";
  htmlBody += "<tr><td style='padding:10px; border:1px solid #ccc;'>Estimated Spend (Cost):</td><td style='padding:10px; border:1px solid #ccc; font-weight:bold; text-align:right; color:#c62828;'>$" + totalInboundCost.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + "</td></tr></table>";

  htmlBody += "<table style='width:100%; border-collapse:collapse;'><tr style='background:#e8f5e9;'><th colspan='2' style='padding:10px; border:1px solid #ccc; color:#2e7d32; font-size:16px;'>📤 Outbound Shipping</th></tr>";
  htmlBody += "<tr><td style='padding:10px; border:1px solid #ccc;'>Total Units Shipped:</td><td style='padding:10px; border:1px solid #ccc; font-weight:bold; text-align:right;'>" + totalOutboundQty + "</td></tr>";
  htmlBody += "<tr><td style='padding:10px; border:1px solid #ccc;'>Estimated Revenue:</td><td style='padding:10px; border:1px solid #ccc; font-weight:bold; text-align:right; color:#2e7d32;'>$" + totalOutboundRev.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + "</td></tr></table>";
  
  htmlBody += "</div>";

  htmlBody += "<h3 style='color:#333; margin-bottom:5px;'>Top Suppliers (By Volume)</h3>";
  htmlBody += "<table style='width:100%; border-collapse:collapse; font-size:12px;'><tr style='background:#f0f0f0;'><th style='padding:6px; border:1px solid #ccc; text-align:left;'>Supplier</th><th style='padding:6px; border:1px solid #ccc; text-align:center;'>Units Received</th><th style='padding:6px; border:1px solid #ccc; text-align:right;'>Spend</th></tr>" + (inHtml || "<tr><td colspan='3' style='text-align:center; padding:10px;'>No inbound activity.</td></tr>") + "</table>";

  htmlBody += "<h3 style='color:#333; margin-top:20px; margin-bottom:5px;'>Top Customers (By Volume)</h3>";
  htmlBody += "<table style='width:100%; border-collapse:collapse; font-size:12px;'><tr style='background:#f0f0f0;'><th style='padding:6px; border:1px solid #ccc; text-align:left;'>Customer</th><th style='padding:6px; border:1px solid #ccc; text-align:center;'>Units Shipped</th><th style='padding:6px; border:1px solid #ccc; text-align:right;'>Revenue</th></tr>" + (outHtml || "<tr><td colspan='3' style='text-align:center; padding:10px;'>No outbound activity.</td></tr>") + "</table>";

  htmlBody += "</div>";

  // Change this array later when you want to add Jessica
  var emailTo = ADMIN_EMAIL; 

  MailApp.sendEmail({
    to: emailTo,
    subject: "📈 ASP Monthly Volume Report - " + today.toLocaleDateString(),
    htmlBody: htmlBody,
    name: COMPANY_NAME,
    attachments: [csvBlob]
  });
}

function testSendOnHandReport() {
  sendAutomatedOnHandReport("Quarterly");
  sendAutomatedOnHandReport("Monthly");
  sendAutomatedOnHandReport("Weekly");
  sendAutomatedOnHandReport("Daily"); 
}

function sendDailyOnHandReport() {
  var dayOfWeek = new Date().getDay();
  
  // 0 is Sunday, 6 is Saturday. Abort the script if it is the weekend.
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return; 
  }
  
  sendAutomatedOnHandReport("Daily"); 
}

function sendWeeklyOnHandReport() {
  sendAutomatedOnHandReport("Weekly");
}

function sendMonthlyOnHandReport() {
  var today = new Date();
  var date = today.getDate();
  var day = today.getDay(); // 0 is Sunday, 6 is Saturday
  
  // Determine if today is the first business day of the month
  var isFirstBusinessDay = (date === 1 && day >= 1 && day <= 5) || // 1st is Mon-Fri
                           (date === 2 && day === 1) ||            // 2nd is Mon (1st was Sun)
                           (date === 3 && day === 1);              // 3rd is Mon (1st was Sat)
  
  if (!isFirstBusinessDay) {
    return; 
  }
  
  sendAutomatedOnHandReport("Monthly");
}

function sendQuarterlyOnHandReport() {
  var today = new Date();
  var month = today.getMonth() + 1; // getMonth is 0-indexed (January is 0)
  var date = today.getDate();
  var day = today.getDay();
  
  // Only proceed during Quarterly kick-off months (Jan, Apr, Jul, Oct)
  if (month !== 1 && month !== 4 && month !== 7 && month !== 10) {
    return;
  }
  
  var isFirstBusinessDay = (date === 1 && day >= 1 && day <= 5) || 
                           (date === 2 && day === 1) || 
                           (date === 3 && day === 1);
                           
  if (!isFirstBusinessDay) {
    return;
  }
  
  sendAutomatedOnHandReport("Quarterly");
}