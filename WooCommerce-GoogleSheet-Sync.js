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
    //Change the headers name in the way you want. BE CAREFUL YOU SHOULD CHANGE IT IN LINE 50 TO 130 
  var headers = ['ID', 'Type', 'SKU', 'نام محصول', 'موجود است؟', 'موجودی', 'قیمت حراجی', 'قیمت عادی', 'Parent', 'Update'];
  if (page == 1 || !sheet) {
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
    products.forEach(function (product) {
      SpreadsheetApp.flush();
      var rowData = [];
      headers.forEach(function (header) {
        switch (header) {
          case 'ID':
            rowData.push(product.id);
            break;
          case 'Type':
            rowData.push(product.type);
            break;
          case 'SKU':
            rowData.push(product.sku);
            break;
          case 'نام محصول':
            rowData.push(product.name);
            break;
          case 'موجود است؟':
            rowData.push(product.stock_status === 'instock' ? 1 : 0);
            break;
          case 'موجودی':
            rowData.push(product.stock_quantity || "");
            break;
          case 'قیمت حراجی':
            rowData.push(product.type === 'variable' ? '' : (product.sale_price || ''));
            break;
          case 'قیمت عادی':
            rowData.push(product.type === 'variable' ? '' : (product.regular_price || product.price || ''));
            break;
          case 'Parent':
            rowData.push('');
            break;
          case 'Update':
            rowData.push(false);
            break;
          default:
            rowData.push('');
        }
      });
      productData.push(rowData);
      // For variable products, fetch and write variations
      if (product.type === 'variable') {
        var variations = fetchVariations(siteUrl, consumerKey, consumerSecret, product.id);
        if (variations) {
          variations.forEach(function (variation) {
            var variationRowData = [];
            headers.forEach(function (header) {
              switch (header) {
                case 'ID':
                  variationRowData.push(variation.id);
                  break;
                case 'Type':
                  variationRowData.push('variation');
                  break;
                case 'SKU':
                  variationRowData.push(variation.sku);
                  break;
                case 'نام محصول':
                  variationRowData.push(product.name + ' - ' + getVariationAttributes(variation));
                  break;
                case 'موجود است؟':
                  variationRowData.push(variation.stock_status === 'instock' ? 1 : 0);
                  break;
                case 'موجودی':
                  variationRowData.push(variation.stock_quantity || "");
                  break;
                case 'قیمت حراجی':
                  variationRowData.push(variation.sale_price || '');
                  break;
                case 'قیمت عادی':
                  variationRowData.push(variation.regular_price || variation.price || '');
                  break;
                case 'Parent':
                  variationRowData.push("id:" + product.id);
                  break;
                case 'Update':
                  variationRowData.push(false);
                  break;
                default:
                  variationRowData.push('');
              }
            });
            productData.push(variationRowData);
          });
        }
      }
      // Auto-resize columns periodically (every 10 products)
      if (currentRow % 10 === 0) {
        sheet.autoResizeColumns(1, headers.length);
      }
    });
    if (productData.length > 0) {
      productData.forEach(function (product) {
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
  hideColumns(sheet, ['ID', 'Type', 'Parent'], headers);

  SpreadsheetApp.getActiveSpreadsheet().toast("Product fetching complete.", 'Success!', 20);

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

// Update product prices
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
    .addItem('Synchronize Prices & Stock', 'updateProductPricesCaller')
    .addItem('Fetch Products', 'fetchProducts')
    .addToUi();
  refreshUpdateColumn();
}

// Refreshes the "Update" column with checkboxes
function refreshUpdateColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var updateColIndex = headers.indexOf("Update") + 1;

  if (updateColIndex > 0) {
    sheet.deleteColumn(updateColIndex);
  }

  lastColumn = sheet.getLastColumn();
  sheet.insertColumnAfter(lastColumn);
  sheet.getRange(1, lastColumn + 1).setValue("Update");

  var dataRange = sheet.getRange(2, lastColumn + 1, sheet.getLastRow() - 1, 1);
  dataRange.insertCheckboxes();
}

// Sets the "Update" checkbox to true when a price column is edited
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  var row = range.getRow();
  var col = range.getColumn();

  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var updateCol = headers.indexOf("Update") + 1;

  var priceColumns = [5, 6, 7, 8];
  if (priceColumns.indexOf(col) !== -1) {
    var checkboxCell = sheet.getRange(row, updateCol);
    checkboxCell.setValue(true);
  }
}

// Updates product prices in WooCommerce based on data in the sheet
function updateProductPrices(consumerKey, consumerSecret, siteUrl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var idCol = 0;
  var typeCol = 1;
  var salePriceCol = 6;
  var regularPriceCol = 7;
  var inStockCol = 4;
  var stockQuantityCol = 5;
  var variationSalePriceCol = 6;
  var variationRegularPriceCol = 7;
  var variationStockQuantityCol = 5;
  var updateCol = data[0].indexOf("Update") + 1;

  var simpleBatch = [];
  var variationBatches = {};
  var batchSize = 100;

  for (var i = 1; i < data.length; i++) {
    var productId = data[i][idCol];
    var productType = data[i][typeCol];
    var salePrice = data[i][salePriceCol];
    var regularPrice = data[i][regularPriceCol];
    var inStock = data[i][inStockCol];
    var stockQuantity = data[i][stockQuantityCol];
    var update = data[i][updateCol - 1];

    Logger.log("Row: " + i);
    Logger.log("Product ID: " + productId);
    Logger.log("Product Type: " + productType);
    Logger.log("Regular Price: " + regularPrice);
    Logger.log("Sale Price: " + salePrice);
    Logger.log("In Stock: " + inStock);
    Logger.log("Stock Quantity: " + stockQuantity);
    Logger.log("Update: " + update);

    if (update === true) {
      var productData = {
        id: productId,
        regular_price: regularPrice ? regularPrice.toString() : "",
        sale_price: salePrice ? salePrice.toString() : "",
        stock_quantity: stockQuantity ? parseInt(stockQuantity) : 0,
        manage_stock: true,
        stock_status: inStock ? 'instock' : 'outofstock'
      };

      if (productType === 'simple') {
        simpleBatch.push(productData);

        if (simpleBatch.length >= batchSize) {
          sendBatchRequest({ update: simpleBatch }, siteUrl + 'wp-json/wc/v3/products/batch', consumerKey, consumerSecret);
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
          regular_price: regularPrice ? regularPrice.toString() : "",
          sale_price: salePrice ? salePrice.toString() : "",
          stock_quantity: parseInt(stockQuantity) ? parseInt(stockQuantity) : 0,
          manage_stock: true,
          stock_status: inStock ? 'instock' : 'outofstock'
        });

        if (variationBatches[parentId].length >= batchSize) {
          sendBatchRequest({ update: variationBatches[parentId] }, siteUrl + 'wp-json/wc/v3/products/' + parentId + '/variations/batch', consumerKey, consumerSecret);
          variationBatches[parentId] = [];
        }
      }

      sheet.getRange(i + 1, updateCol).setValue(false);
    }
  }

  if (simpleBatch.length > 0) {
    sendBatchRequest({ update: simpleBatch }, siteUrl + 'wp-json/wc/v3/products/batch', consumerKey, consumerSecret);
  }

  for (var parentId in variationBatches) {
    if (variationBatches[parentId].length > 0) {
      sendBatchRequest({ update: variationBatches[parentId] }, siteUrl + 'wp-json/wc/v3/products/' + parentId + '/variations/batch', consumerKey, consumerSecret);
    }
  }
}

// Finds the parent product ID for a variation
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

  return null;
}

// Sends a batch request to the WooCommerce API
function sendBatchRequest(payload, url, consumerKey, consumerSecret) {
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
    Logger.log("Response Content: " + response.getContentText());
  } catch (error) {
    Logger.log("Error in batch update: " + error.message);
    Logger.log("Request Payload: " + JSON.stringify(payload));
    Logger.log("Response Code: " + error.responseCode);
    Logger.log("Response Headers: " + JSON.stringify(error.getHeaders()));
  }
}

