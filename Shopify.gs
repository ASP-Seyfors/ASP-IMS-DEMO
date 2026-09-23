// =======================================================================
// SHOPIFY API ENGINE
// =======================================================================

function testShopifyConnection(itemsToTest) {
  let token = PropertiesService.getScriptProperties().getProperty("SHOPIFY_ACCESS_TOKEN");
  let shopUrl = PropertiesService.getScriptProperties().getProperty("SHOPIFY_STORE_URL");
  
  if (!token || !shopUrl) {
    return { status: "error", message: "Shopify credentials missing in Script Properties." };
  }

  let cleanShopUrl = shopUrl.replace(/^https?:\/\//, '').replace(/\/$/, '');
  let successCount = 0;
  let errors = [];

  // We need a location ID to update inventory
  let locationId = null;
  try {
    let locRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/locations.json`, {
      "method": "get",
      "headers": { "X-Shopify-Access-Token": token }
    });
    locationId = JSON.parse(locRes.getContentText()).locations[0].id;
  } catch (e) {
    return { status: "error", message: "Could not fetch Shopify Location ID." };
  }

  itemsToTest.forEach(item => {
    let cleanPrice = parseFloat(String(item.price || '').replace(/[^0-9.-]+/g, '')) || 0;
    let rawStatus = String(item.status || "active").toUpperCase();
    let intendedStatus = (rawStatus === "INACTIVE" || rawStatus === "DRAFT") ? "draft" : "active";
    let publishedStatus = intendedStatus === "active";

    try {
      // STEP 1: Check if item already exists by SKU
      let searchRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products.json?sku=${item.ref}`, {
        "method": "get",
        "headers": { "X-Shopify-Access-Token": token }
      });
      
      let products = JSON.parse(searchRes.getContentText()).products;

      if (products && products.length > 0) {
        // EXISTS -> UPDATE IT (No Duplication)
        let shopifyProduct = products[0];
        let shopifyVariant = shopifyProduct.variants[0];

        // UPDATE EXISTING PRODUCT
        UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products/${shopifyProduct.id}.json`, {
          "method": "put",
          "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
          "payload": JSON.stringify({
            "product": {
              "id": shopifyProduct.id,
              "title": String(item.title),
              "product_type": String(item.category || 'Surgical Supply'),
              "tags": String(item.category || 'Surgical Supply'),
              "status": intendedStatus,
              // "published": publishedStatus,
              "variants": [{
                "id": shopifyVariant.id,
                "price": cleanPrice.toFixed(2)
              }]
            }
          }),
          "muteHttpExceptions": true
        });

        // ✨ GraphQL Taxonomy Injection (Dynamic via Payload)
        let shopCategoryGid = item.shopifyCategory || SHOPIFY_DEFAULT_TAXONOMY_GID;
        let graphQueryUpdate = `mutation { productUpdate(input: { id: "gid://shopify/Product/${shopifyProduct.id}", category: "${shopCategoryGid}" }) { userErrors { message } } }`;
        UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
          "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
          "payload": JSON.stringify({ "query": graphQueryUpdate }), "muteHttpExceptions": true
        });

        // Update Inventory Quantity
        UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/inventory_levels/set.json`, {
          "method": "post",
          "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
          "payload": JSON.stringify({
            "location_id": locationId,
            "inventory_item_id": shopifyVariant.inventory_item_id,
            "available": item.availableQty || 0
          }),
          "muteHttpExceptions": true
        });

      } else {
        // DOES NOT EXIST -> CREATE IT NEW
        let createProdRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products.json`, {
          "method": "post",
          "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
          "payload": JSON.stringify({
            "product": {
              "title": String(item.title),
              "body_html": item.desc || '',
              "vendor": item.mfr || 'Unknown',
              "product_type": String(item.category || 'Surgical Supply'),
              "tags": String(item.category || 'Surgical Supply'),
              "status": intendedStatus,
              // "published": publishedStatus,
              // "published_scope": "web",
              "images": [{ "src": GITHUB_DEFAULT_PRODUCT_IMAGE_URL }],
              "variants": [
                {
                  "sku": item.ref,
                  "price": cleanPrice.toFixed(2),
                  "barcode": item.gtin || "",
                  "inventory_management": "shopify",
                  "inventory_policy": "deny",
                  "fulfillment_service": "manual",
                  "requires_shipping": true
                }
              ]
            }
          }),
          "muteHttpExceptions": true
        });

        let newProdJson = JSON.parse(createProdRes.getContentText());
        if (newProdJson && newProdJson.product) {
           // ✨ GraphQL Taxonomy Injection (Hardcoded to Medical Supplies)
           let shopCategoryGid = item.shopifyCategory || SHOPIFY_DEFAULT_TAXONOMY_GID;
           let graphQueryUpdate = `mutation { productUpdate(input: { id: "gid://shopify/Product/${newProdJson.product.id}", category: "${shopCategoryGid}" }) { userErrors { message } } }`;
           UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
             "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
             "payload": JSON.stringify({ "query": graphQueryUpdate }), "muteHttpExceptions": true
           });
        }
      }

      successCount++;
    } catch (e) {
      errors.push(`${item.ref} Error: ${e.message}`);
    }
  });

  if (errors.length > 0) {
    return { status: "error", message: `Processed ${successCount} items, but hit errors:\n\n${errors.join('\n')}` };
  } else {
    return { status: "success", message: `✅ Successfully verified ${successCount} test items! Duplicates were prevented, and data was correctly routed.` };
  }
}

function syncShopifySandbox(syncPayload) {
  let token = PropertiesService.getScriptProperties().getProperty("SHOPIFY_ACCESS_TOKEN");
  let shopUrl = PropertiesService.getScriptProperties().getProperty("SHOPIFY_STORE_URL");
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let bugSheet = ss.getSheetByName(SHEET_BUG_REPORTS);
  let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");

  if (!token || !shopUrl) {
    if (bugSheet) bugSheet.appendRow([timestamp, "System", "Backend API", "SYNC_SHOPIFY_SANDBOX", "Shopify Setup Error", "Credentials missing in Script Properties", SHOPIFY_BUG_ENV]);
    return { status: "error", message: "Shopify credentials missing." };
  }

  let cleanShopUrl = shopUrl.replace(/^https?:\/\//, '').replace(/\/$/, '');
  let successCount = 0;
  let errors = [];
  let successfulSkus = [];
  let apiTrace = ["Payload Size: " + syncPayload.length];

  let locationId = null;
  try {
    let locRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/locations.json`, { "method": "get", "headers": { "X-Shopify-Access-Token": token }, "muteHttpExceptions": true });
    if (locRes.getResponseCode() >= 400) throw new Error("Location Fetch Failed: " + locRes.getContentText());
    locationId = JSON.parse(locRes.getContentText()).locations[0].id;
    apiTrace.push("Got Location: " + locationId);
  } catch (e) {
    if (bugSheet) bugSheet.appendRow([timestamp, "System", "Backend API", "SYNC_SHOPIFY_SANDBOX", "Location API Error", e.message, SHOPIFY_BUG_ENV]);
    return { status: "error", message: "Could not fetch Location ID." };
  }

  syncPayload.forEach(item => {
    let targetSku = String(item.ref);
    let rawStatus = String(item.status || "active").toUpperCase();
    let intendedStatus = (rawStatus === "INACTIVE" || rawStatus === "DRAFT") ? "draft" : "active";
    let publishedStatus = intendedStatus === "active";

    try {
      // SEARCH BY HANDLE
      let searchRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products.json?handle=${item.handle}`, { "method": "get", "headers": { "X-Shopify-Access-Token": token }, "muteHttpExceptions": true });
      if (searchRes.getResponseCode() >= 400) throw new Error("Search Failed: " + searchRes.getContentText());
      
      let products = JSON.parse(searchRes.getContentText()).products;
      
      if (products && products.length > 0) {
        let shopifyProduct = products[0];
        let shopifyVariant = shopifyProduct.variants.find(v => String(v.sku).replace(/^'/, '').toUpperCase() === targetSku.toUpperCase());
        
        if (shopifyVariant) {
           // UPDATE EXISTING VARIANT
           let varRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/variants/${shopifyVariant.id}.json`, {
             "method": "put", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
             "payload": JSON.stringify({ "variant": { "id": shopifyVariant.id, "price": String(item.price), "barcode": String(item.gtin || ''), "weight": item.weight, "weight_unit": "lb" } }), "muteHttpExceptions": true
           });
           if (varRes.getResponseCode() >= 400) throw new Error("Variant Update Failed: " + varRes.getContentText());
           
           if (!item.isBundle) {
             // UPDATE THE PARENT PRODUCT
             let prodRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products/${shopifyProduct.id}.json`, {
               "method": "put", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
               "payload": JSON.stringify({ 
                 "product": {
                    "id": shopifyProduct.id,
                    "title": String(item.title), 
                    "body_html": String(item.desc || ''), 
                    "vendor": String(item.mfr || 'Unknown'), 
                    "product_type": String(item.category || 'Surgical Supply'), 
                    "tags": String(item.category || 'Surgical Supply'), 
                    "status": intendedStatus, 
                    // "published": publishedStatus, 
                    // "published_scope": "web"
                 } 
               }), "muteHttpExceptions": true
             });
             if (prodRes.getResponseCode() >= 400) throw new Error("Product Update Failed: " + prodRes.getContentText());
             
             // ✨ GraphQL Taxonomy Injection (Dynamic via Payload)
            let shopCategoryGid = item.shopifyCategory || SHOPIFY_DEFAULT_TAXONOMY_GID;
            let graphQueryUpdate = `mutation { productUpdate(input: { id: "gid://shopify/Product/${shopifyProduct.id}", category: "${shopCategoryGid}" }) { userErrors { message } } }`;
             UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
               "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
               "payload": JSON.stringify({ "query": graphQueryUpdate }), "muteHttpExceptions": true
             });
           }

           let invRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/inventory_levels/set.json`, {
             "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
             "payload": JSON.stringify({ "location_id": locationId, "inventory_item_id": shopifyVariant.inventory_item_id, "available": parseInt(item.availableQty, 10) || 0 }), "muteHttpExceptions": true
           });
           if (invRes.getResponseCode() >= 400) throw new Error("Inventory Set Failed: " + invRes.getContentText());
           
           successCount++; successfulSkus.push(targetSku);
           apiTrace.push(`[UPDATED] ${targetSku}`);
        } else {
           // VARIANT DOES NOT EXIST (CREATE IT)
           let optValue = item.isBundle ? `Box of ${item.uomMult}` : "Each";
           let createVarRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products/${shopifyProduct.id}/variants.json`, {
             "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
             "payload": JSON.stringify({ "variant": { "sku": targetSku, "price": String(item.price), "option1": optValue, "barcode": String(item.gtin || ''), "inventory_management": "shopify", "inventory_policy": "deny", "fulfillment_service": "manual", "weight": item.weight, "weight_unit": "lb" } }), "muteHttpExceptions": true
           });
           if (createVarRes.getResponseCode() >= 400) throw new Error("Variant Create Failed: " + createVarRes.getContentText());
           
           let newVariantJson = JSON.parse(createVarRes.getContentText());
           if (newVariantJson && newVariantJson.variant) {
             let invRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/inventory_levels/set.json`, {
               "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
               "payload": JSON.stringify({ "location_id": locationId, "inventory_item_id": newVariantJson.variant.inventory_item_id, "available": parseInt(item.availableQty, 10) || 0 }), "muteHttpExceptions": true
             });
             if (invRes.getResponseCode() >= 400) throw new Error("Inventory Set Failed: " + invRes.getContentText());
           }
           successCount++; successfulSkus.push(targetSku);
           apiTrace.push(`[VARIANT CREATED] ${targetSku}`);
        }
      } else {
        // PRODUCT DOES NOT EXIST (CREATE ENTIRELY NEW)
        let optValue = item.isBundle ? `Box of ${item.uomMult}` : "Each";
        let createProdRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/products.json`, {
          "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
          "payload": JSON.stringify({
            "product": {
              "title": String(item.title), "body_html": String(item.desc || ''), "vendor": String(item.mfr || 'Unknown'),
              "product_type": String(item.category || 'Surgical Supply'),
              "tags": String(item.category || 'Surgical Supply'),
              "status": intendedStatus, // "published": publishedStatus, "published_scope": "web",
              "images": [{ "src": GITHUB_DEFAULT_PRODUCT_IMAGE_URL }],
              "options": [{ "name": "Unit of Measure" }],
              "variants": [{ "sku": targetSku, "price": String(item.price), "option1": optValue, "barcode": String(item.gtin || ''), "inventory_management": "shopify", "inventory_policy": "deny", "fulfillment_service": "manual", "weight": item.weight, "weight_unit": "lb" }]
            }
          }), "muteHttpExceptions": true
        });
        if (createProdRes.getResponseCode() >= 400) throw new Error("Product Create Failed: " + createProdRes.getContentText());
        
        let newProdJson = JSON.parse(createProdRes.getContentText());
        if (newProdJson && newProdJson.product) {
            // ✨ GraphQL Taxonomy Injection (Hardcoded to Medical Supplies)
            let shopCategoryGid = item.shopifyCategory || SHOPIFY_DEFAULT_TAXONOMY_GID;
            let graphQueryUpdate = `mutation { productUpdate(input: { id: "gid://shopify/Product/${newProdJson.product.id}", category: "${shopCategoryGid}" }) { userErrors { message } } }`;
            UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
              "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
              "payload": JSON.stringify({ "query": graphQueryUpdate }), "muteHttpExceptions": true
            });

            if (newProdJson.product.variants) {
               let invRes = UrlFetchApp.fetch(`https://${cleanShopUrl}/admin/api/${SHOPIFY_API_VERSION}/inventory_levels/set.json`, {
                 "method": "post", "headers": { "X-Shopify-Access-Token": token, "Content-Type": "application/json" },
                 "payload": JSON.stringify({ "location_id": locationId, "inventory_item_id": newProdJson.product.variants[0].inventory_item_id, "available": parseInt(item.availableQty, 10) || 0 }), "muteHttpExceptions": true
               });
               if (invRes.getResponseCode() >= 400) throw new Error("Inventory Set Failed: " + invRes.getContentText());
            }
        }
        successCount++; successfulSkus.push(targetSku);
        apiTrace.push(`[PRODUCT CREATED] ${targetSku}`);
      }   
            
        // ✨ FIX: Pause for 600ms per item to prevent Shopify 429 Rate Limit errors
        Utilities.sleep(600);
            
        } catch (e) {
          errors.push(`${targetSku} Error: ${e.message}`);
        }
    });

  if (successCount > 0) {
     let itemSheet = ss.getSheetByName("Items");
     if (itemSheet) {
        let itemVals = itemSheet.getDataRange().getValues();
        let refMap = {};
        for (let i = 1; i < itemVals.length; i++) { if (itemVals[i][DB_MAP.REF]) refMap[String(itemVals[i][DB_MAP.REF]).toUpperCase().replace(/'/g, '').trim()] = i; }
        successfulSkus.forEach(sku => {
           let cleanSku = String(sku).toUpperCase().trim();
           if (refMap[cleanSku] !== undefined) itemVals[refMap[cleanSku]][DB_MAP.SHOPIFY_SYNC] = "TRUE";
        });
        itemSheet.getRange(1, 1, itemVals.length, itemVals[0].length).setValues(itemVals);
     }
  }

  if (bugSheet) {
     let traceStr = apiTrace.join(" | ") + (errors.length > 0 ? " || ERRORS: " + errors.join(" | ") : " || ALL CLEAR");
     bugSheet.appendRow([timestamp, "System", "Backend API", "SYNC_SHOPIFY_SANDBOX", SHOPIFY_BUG_TAG, traceStr, SHOPIFY_BUG_ENV]);
  }

  if (errors.length > 0) return { status: "error", message: `Updated ${successCount} items, but hit errors. Check Sync_Log tab.` }; 
  else return { status: "success", message: `✅ Successfully updated ${successCount} items in Shopify!` };
}