// Configure WooCommerce API access information
function constants() {
    var consumerKey = 'YOUR_CONSUMER_KEY';
    var consumerSecret = 'YOUR_CONSUMER_SECRET';
    var siteUrl = 'YOUR_SITE_URL'; // e.g., 'https://yourstore.com/'
    return [consumerKey, consumerSecret, siteUrl];
  }
  
  
  function fetchProducts() {
    var consumerKey = constants()[0];
    var consumerSecret = constants()[1];
    var siteUrl = constants()[2];
    siteUrl = siteUrl.replace(/\/$/, '');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tempFetchSheet = ss.getSheetByName("TempFetch");
    if (!tempFetchSheet) {
      tempFetchSheet = ss.insertSheet("TempFetch");
      var page = 1;
      tempFetchSheet.getRange("A1").setValue(page);
    } else {
      var pageValue = tempFetchSheet.getRange("A1").getValue();
      if (typeof pageValue === 'number') {
        var page = pageValue;
      } else {
        var page = 1;
      }
    }
    var sheet = ss.getSheetByName('Products');
    var headers = ['ID', 'Type', 'SKU', 'Name', 'In stock?', 'Stock', 'Sale price', 'Regular price', 'Parent'];
    if ( page == 1 || !sheet )
    {
          if (sheet) {
          ss.deleteSheet(sheet);
          }
          sheet = ss.insertSheet('Products');
        if (sheet.getLastRow() === 0) {
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        }
        page = 1;
    }
    var perPage = 10; // WooCommerce default max per page
    var currentRow = Math.max(sheet.getLastRow() + 1, 2); // Start after headers or last row
    while (true) {
      var products = fetchProductBatch(siteUrl, consumerKey, consumerSecret, page, perPage);
      var productData = [];
      if (!products || products.length === 0) {
        break;
      }
      products.forEach(function(product) {
          SpreadsheetApp.flush();
        productData.push([
          product.id,
          product.type,
          product.sku,
          product.name,
          product.stock_status === 'instock' ? 1 : 0,
          product.stock_quantity || "",
          product.type === 'variable' ? '' : (product.sale_price || ''),
          product.type === 'variable' ? '' : (product.regular_price || product.price || ''),
          ''
        ]);
        // For variable products, fetch and write variations
        if (product.type === 'variable') {
          var variations = fetchVariations(siteUrl, consumerKey, consumerSecret, product.id);
          if (variations) {
            variations.forEach(function(variation) {
              productData.push( [
                variation.id,
                'variation',
                variation.sku,
                product.name + ' - ' + getVariationAttributes(variation),
                variation.stock_status === 'instock' ? 1 : 0,
                variation.stock_quantity || "",
                variation.sale_price || '',
                variation.regular_price || variation.price || '',
                "id:" + product.id
              ]);
            });
          }
        }
        // Auto-resize columns periodically (every 10 products)
        if (currentRow % 10 === 0) {
          sheet.autoResizeColumns(1, headers.length);
        }
      });
      if (productData.length > 0) {
          productData.forEach(function(product) {
              sheet.getRange(currentRow, 1, 1, headers.length).setValues([product]);
              currentRow++;
          });
        }
      page++;
      tempFetchSheet.getRange("A1").setValue(page);
      Utilities.sleep(500);
    }
    // Final auto-resize and hide columns
    sheet.autoResizeColumns(1, headers.length);
    hideColumns(sheet, ['Type', 'SKU', 'Parent'], headers);
    
      SpreadsheetApp.getActiveSpreadsheet().toast("Product fetching complete.", 'Success!',10);
  
    if (tempFetchSheet) {
      ss.deleteSheet(tempFetchSheet);
      }
    // add update check box
    refreshUpdateColumn();
  }
  
  // Get the list of variations for a product
  function fetchVariations(siteUrl, consumerKey, consumerSecret, productId) {
    try {
      var response = makeAuthenticatedRequest(siteUrl + `/wp-json/wc/v3/products/${productId}/variations`, { 'per_page': 100 }, consumerKey, consumerSecret);
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('Error fetching variations for product ' + productId + ': ' + error);
      return null;
    }
  }
  
  // Combine variation attributes as a name
  function getVariationAttributes(variation) {
    return variation.attributes?.map(attr => attr.option).join(', ') || '';
  }
  
  // Get a batch of products from the WooCommerce API
  function fetchProductBatch(siteUrl, consumerKey, consumerSecret, page, perPage) {
    try {
      var response = makeAuthenticatedRequest(siteUrl + '/wp-json/wc/v3/products', { 'per_page': perPage, 'page': page }, consumerKey, consumerSecret);
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('Error fetching products: ' + error);
      return null;
    }
  }
  
  // Send an authenticated request to the API
  function makeAuthenticatedRequest(url, params, consumerKey, consumerSecret) {
    var queryParams = Object.entries(params).map(([key, val]) => `${key}=${encodeURIComponent(val)}`).join('&');
    var options = {
      'method': 'GET',
      'headers': {
        "Authorization": "Basic " + Utilities.base64Encode(consumerKey + ":" + consumerSecret)
      },
      'muteHttpExceptions': true
    };
    return UrlFetchApp.fetch(`${url}?${queryParams}`, options);
  }
  
  // Hide the specified columns in the sheet
  function hideColumns(sheet, columnsToHide, headers) {
    columnsToHide.forEach(columnName => {
      var columnIndex = headers.indexOf(columnName) + 1;
      if (columnIndex > 0) sheet.hideColumns(columnIndex);
    });
  }
  
  // Create an "Update" checkbox column to specify products for price update
  function refreshUpdateColumn() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
    if (!sheet) return;
  
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var updateColIndex = headers.indexOf("Update") + 1;
  
    if (updateColIndex > 0) sheet.deleteColumn(updateColIndex);
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue("Update");
  
    var dataRange = sheet.getRange(2, sheet.getLastColumn(), sheet.getLastRow() - 1, 1);
    dataRange.insertCheckboxes();
  }
  
  // Check for changes in price and mark the update checkbox
  function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== "Products") return;
  
    var row = e.range.getRow();
    var col = e.range.getColumn();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var updateCol = headers.indexOf("Update") + 1;
  
    var priceColumns = [6, 7]; // Price columns (Sale Price, Regular Price)
    if (priceColumns.includes(col)) {
      sheet.getRange(row, updateCol).setValue(true);
    }
  }
  
  function updateProductPricesCaller() {
    var consumerKey = constants()[0];
    var consumerSecret = constants()[1];
    var siteUrl = constants()[2];
    updateProductPrices(consumerKey, consumerSecret, siteUrl);
  }
  
  // Add a custom menu to Google Sheets
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('WordPress')
      .addItem('Send Updated Prices', 'updateProductPrices')
      .addItem('Fetch Products', 'fetchProducts')
  .addToUi();
    refreshUpdateColumn();
  }
  
  function refreshUpdateColumn() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var updateColIndex = headers.indexOf("Update") + 1; // Get the index of the Update column (1-based index)
  
    // If the Update column exists, delete it
    if (updateColIndex > 0) {
      sheet.deleteColumn(updateColIndex);
    }
  
    // Add the new Update column with checkboxes
    lastColumn = sheet.getLastColumn(); // Update lastColumn after potential deletion
    sheet.insertColumnAfter(lastColumn);
    sheet.getRange(1, lastColumn + 1).setValue("Update");
  
    // Add checkboxes for each row in the new "Update" column
    var dataRange = sheet.getRange(2, lastColumn + 1, sheet.getLastRow() - 1, 1);
    dataRange.insertCheckboxes();
  }
  
  function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
  
    // Get the edited cell's row and column
    var row = range.getRow();
    var col = range.getColumn();
  
    // Get the last column and the position of the "Update" column
    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var updateCol = headers.indexOf("Update") + 1; // Add 1 because headers is zero-indexed
  
    // Check if the edited cell is in the Sale Price or Regular Price column
    var priceColumns = [4, 5, 6, 7]; // Adjust based on the actual column positions
    if (priceColumns.indexOf(col) !== -1) {
      // Get the checkbox cell
      var checkboxCell = sheet.getRange(row, updateCol);
  
      // Check the checkbox
      checkboxCell.setValue(true);
    }
  }
  
  function updateProductPrices(consumerKey, consumerSecret, siteUrl) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
  
    var idCol = 0;
    var typeCol = 1;
    var regularPriceCol = 5;
    var salePriceCol = 4;
    var variationRegularPriceCol = 5;
    var variationSalePriceCol = 4;
    var updateCol = data[0].indexOf("Update"); // Get the actual column index of "Update"
  
    var simpleBatch = [];
    var variationBatches = {};
    var batchSize = 100; // Maximum number of products to update in one batch
  
    for (var i = 1; i < data.length; i++) { // Start from 1 to skip the header row
      var productId = data[i][idCol];
      var productType = data[i][typeCol];
      var regularPrice = data[i][regularPriceCol];
      var salePrice = data[i][salePriceCol];
      var variationRegularPrice = data[i][variationRegularPriceCol];
      var variationSalePrice = data[i][variationSalePriceCol];
      var update = data[i][updateCol];
  
      if (update) {
        if (productType === 'simple') {
          simpleBatch.push({
            id: productId,
            regular_price: regularPrice ? regularPrice.toString() : "",
            sale_price: salePrice ? salePrice.toString() : ""
          });
  
          if (simpleBatch.length >= batchSize) {
            sendBatchRequest(simpleBatch, siteUrl + 'wp-json/wc/v3/products/batch', consumerKey, consumerSecret);
            simpleBatch = [];
          }
        } else if (productType === 'variation') {
          var variationId = productId;
          var parentId = findParentProductId(data, variationId, idCol, typeCol);
  
          if (!variationBatches[parentId]) {
            variationBatches[parentId] = [];
          }
  
          variationBatches[parentId].push({
            id: variationId,
            regular_price: variationRegularPrice ? variationRegularPrice.toString() : "",
            sale_price: variationSalePrice ? variationSalePrice.toString() : ""
          });
  
          if (variationBatches[parentId].length >= batchSize) {
            sendBatchRequest(variationBatches[parentId], siteUrl + 'wp-json/wc/v3/products/' + parentId + '/variations/batch', consumerKey, consumerSecret);
            variationBatches[parentId] = [];
          }
        }
  
        sheet.getRange(i + 1, updateCol + 1).setValue(false); // Uncheck the checkbox after updating
      }
    }
  
    // Send any remaining requests in the batches
    if (simpleBatch.length > 0) {
      sendBatchRequest(simpleBatch, siteUrl + 'wp-json/wc/v3/products/batch', consumerKey, consumerSecret);
    }
  
    for (var parentId in variationBatches) {
      if (variationBatches[parentId].length > 0) {
        sendBatchRequest(variationBatches[parentId], siteUrl + 'wp-json/wc/v3/products/' + parentId + '/variations/batch', consumerKey, consumerSecret);
      }
    }
  }
  
  function findParentProductId(data, variationId, idCol, typeCol) {
    var parentId = variationId - 1;
  
    while (parentId > 0) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][idCol] == parentId && data[i][typeCol] == 'variable') {
          return parentId;
        }
      }
      parentId--;
    }
  
    return null; // or throw an error if the parent product ID is not found
  }
  
  function sendBatchRequest(batch, url, consumerKey, consumerSecret) {
    var payload = {
      update: batch
    };
  
    var options = {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: {
        "Authorization": "Basic " + Utilities.base64Encode(consumerKey + ":" + consumerSecret)
      }
    };
  
    try {
      var response = UrlFetchApp.fetch(url, options);
      Logger.log("Batch update successful: " + response.getContentText());
    } catch (error) {
      Logger.log("Error in batch update: " + error.message);
    }
  }

