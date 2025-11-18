// Configure WooCommerce API access information

/**
 * پیکربندی اطلاعات دسترسی به API ووکامرس.
 * @return {string[]} [ConsumerKey, ConsumerSecret, SiteUrl]
 */
function constants() {
  var consumerKey = 'YOUR_CONSUMER_KEY';
  var consumerSecret = 'YOUR_CONSUMER_SECRET';
  // Don't use / at the end of the URL
  var siteUrl = 'YOUR_SITE_URL'; // e.g., 'https://yourstore.com'
  return [consumerKey, consumerSecret, siteUrl];
}

/**
 * عنوان‌های ستون‌ها در شیت "Products".
 */
const HEADERS = ['ID', 'Type', 'SKU', 'نام محصول', 'موجود است؟', 'موجودی', 'قیمت حراجی', 'قیمت عادی', 'Parent', 'Update'];
const MAX_PRODUCTS_PER_RUN = 500;
// --- توابع کمکی (Helper Functions) ---

/**
 * ارسال یک درخواست احراز هویت شده به API ووکامرس.
 * @param {string} url آدرس کامل API
 * @param {object} params پارامترهای Query
 * @param {string} consumerKey کلید مصرف کننده
 * @param {string} consumerSecret رمز مصرف کننده
 * @return {HTTPResponse|null} پاسخ HTTP یا null در صورت خطا
 */
function makeAuthenticatedRequest(url, params, consumerKey, consumerSecret, method = 'GET', payload = null) {
  const queryParams = Object.entries(params).map(([key, val]) => `${key}=${encodeURIComponent(val)}`).join('&');
  const options = {
    method: method,
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(consumerKey + ":" + consumerSecret)
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }

  try {
    const response = UrlFetchApp.fetch(`${url}?${queryParams}`, options);
    // مدیریت خطاهای API (مثل Rate Limiting)
    if (response.getResponseCode() >= 400) {
      Logger.log(`API Error: ${response.getResponseCode()} - ${response.getContentText()}`);
      return null;
    }
    return response;
  } catch (error) {
    Logger.log(`Fetch Error: ${error.message} on URL ${url}`);
    return null;
  }
}

/**
 * پنهان کردن ستون‌های مشخص شده در شیت.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet شیت مورد نظر
 * @param {string[]} columnsToHide نام ستون‌های برای پنهان شدن
 * @param {string[]} headers هدرهای ستون‌ها
 */
function hideColumns(sheet, columnsToHide, headers) {
  columnsToHide.forEach(columnName => {
    const columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex > 0) sheet.hideColumns(columnIndex);
  });
}

/**
 * ترکیب مشخصه‌های ورییشن به عنوان یک نام.
 * @param {object} variation شیء ورییشن
 * @return {string} رشته ترکیبی
 */
function getVariationAttributes(variation) {
  return variation.attributes?.map(attr => attr.option).join(', ') || '';
}

// --- توابع واکشی (Fetch Functions) ---

/**
 * واکشی ورییشن‌های یک محصول مادر از WooCommerce API.
 * @param {string} siteUrl آدرس سایت
 * @param {string} consumerKey کلید
 * @param {string} consumerSecret رمز
 * @param {number} productId ID محصول مادر
 * @return {object[]|null} آرایه ورییشن‌ها یا null
 */
function fetchVariations(siteUrl, consumerKey, consumerSecret, productId) {
  const url = `${siteUrl}/wp-json/wc/v3/products/${productId}/variations`;
  // واکشی حداکثر 100 ورییشن در یک درخواست
  const response = makeAuthenticatedRequest(url, { 'per_page': 100 }, consumerKey, consumerSecret);

  if (!response) return null;

  try {
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Error parsing variations for product ${productId}: ${error}`);
    return null;
  }
}

/**
 * واکشی محصولات از WooCommerce API (حداکثر 100 محصول در هر صفحه).
 * @param {string} siteUrl آدرس سایت
 * @param {string} consumerKey کلید
 * @param {string} consumerSecret رمز
 * @param {number} page شماره صفحه
 * @return {object[]|null} آرایه محصولات یا null
 */
function fetchProductBatch(siteUrl, consumerKey, consumerSecret, page) {
  const perPage = 100; // حداکثر مجاز ووکامرس
  const url = `${siteUrl}/wp-json/wc/v3/products`;
  // استفاده از status: publish برای واکشی فقط محصولات منتشر شده
  const response = makeAuthenticatedRequest(url, { 'per_page': perPage, 'page': page, 'status': 'publish' }, consumerKey, consumerSecret);

  if (!response) return null;

  try {
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Error parsing products on page ${page}: ${error}`);
    return null;
  }
}

/**
 * واکشی محصولات و ورییشن‌ها به صورت بهینه و صفحه‌بندی شده.
 */
function fetchProducts() {
  const [consumerKey, consumerSecret, siteUrl] = constants();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- مدیریت شیت موقت برای Continuation ---
  let tempFetchSheet = ss.getSheetByName("TempFetch");
  if (!tempFetchSheet) tempFetchSheet = ss.insertSheet("TempFetch");
  let currentPage = tempFetchSheet.getRange("A1").getValue() || 1;
  if (typeof currentPage !== 'number' || currentPage < 1) currentPage = 1;

  // --- مدیریت شیت اصلی محصولات ---
  let sheet = ss.getSheetByName('Products');
  if (currentPage === 1) {
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet('Products');
    // نوشتن هدرها فقط یکبار
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }

  let productDataToAppend = [];
  let currentRow = Math.max(sheet.getLastRow() + 1, 2);
  let productsFetchedCount = 0;

  SpreadsheetApp.getActiveSpreadsheet().toast(`شروع واکشی از صفحه ${currentPage}...`, 'عملیات در حال انجام', -1);
  const startTime = new Date();

  while (productsFetchedCount < MAX_PRODUCTS_PER_RUN) {
    const products = fetchProductBatch(siteUrl, consumerKey, consumerSecret, currentPage);

    if (!products || products.length === 0) {
      Logger.log(`No more products found on page ${currentPage}.`);
      break; // پایان واکشی
    }

    for (const product of products) {
      if (productsFetchedCount >= MAX_PRODUCTS_PER_RUN) break;

      // --- افزودن محصول اصلی ---
      const productRow = HEADERS.map(header => {
        switch (header) {
          case 'ID': return product.id;
          case 'Type': return product.type;
          case 'SKU': return product.sku;
          case 'نام محصول': return product.name;
          case 'موجود است؟': return product.stock_status === 'instock' ? 1 : 0;
          case 'موجودی': return product.stock_quantity || '';
          case 'قیمت حراجی': return product.type === 'variable' ? '' : (product.sale_price || '');
          case 'قیمت عادی': return product.type === 'variable' ? '' : (product.regular_price || product.price || '');
          case 'Parent': return '';
          case 'Update': return false;
          default: return '';
        }
      });
      productDataToAppend.push(productRow);
      productsFetchedCount++;

      // --- افزودن ورییشن‌ها (فقط برای محصول متغیر) ---
      if (product.type === 'variable') {
        const variations = fetchVariations(siteUrl, consumerKey, consumerSecret, product.id);
        if (variations) {
          for (const variation of variations) {
            if (productsFetchedCount >= MAX_PRODUCTS_PER_RUN) break;

            const variationRow = HEADERS.map(header => {
              switch (header) {
                case 'ID': return variation.id;
                case 'Type': return 'variation';
                case 'SKU': return variation.sku;
                case 'نام محصول': return `${product.name} - ${getVariationAttributes(variation)}`;
                case 'موجود است؟': return variation.stock_status === 'instock' ? 1 : 0;
                case 'موجودی': return variation.stock_quantity || '';
                case 'قیمت حراجی': return variation.sale_price || '';
                case 'قیمت عادی': return variation.regular_price || variation.price || '';
                case 'Parent': return `id:${product.id}`; // فرمت استاندارد برای تشخیص والد
                case 'Update': return false;
                default: return '';
              }
            });
            productDataToAppend.push(variationRow);
            productsFetchedCount++;
          }
        }
      }
    }

    // در پایان هر صفحه، داده‌ها را یکجا بنویسید
    if (productDataToAppend.length > 0) {
      sheet.getRange(currentRow, 1, productDataToAppend.length, HEADERS.length).setValues(productDataToAppend);
      currentRow += productDataToAppend.length;
      productDataToAppend = []; // پاک کردن بافر
    }

    // اگر از محدودیت زمان اجرا نزدیک شد، متوقف کنید
    if ((new Date().getTime() - startTime.getTime()) > 300000) { // 5 دقیقه
      Logger.log("Nearing execution limit. Stopping fetch.");
      currentPage++;
      tempFetchSheet.getRange("A1").setValue(currentPage);
      SpreadsheetApp.getActiveSpreadsheet().toast(`واکشی موقت متوقف شد. از صفحه ${currentPage} ادامه خواهد یافت.`, 'هشدار زمان اجرا', 20);
      return;
    }

    currentPage++;
    tempFetchSheet.getRange("A1").setValue(currentPage);
    // تأخیر ملایم بین صفحات برای احترام به Rate Limit ووکامرس
    Utilities.sleep(100);
  }

  // --- نهایی‌سازی ---
  if (productDataToAppend.length > 0) {
    sheet.getRange(currentRow, 1, productDataToAppend.length, HEADERS.length).setValues(productDataToAppend);
  }

  sheet.autoResizeColumns(1, HEADERS.length);
  hideColumns(sheet, ['ID', 'Type', 'Parent'], HEADERS);

  // حذف شیت موقت و نمایش پیام تکمیل
  ss.deleteSheet(tempFetchSheet);
  refreshUpdateColumn(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("واکشی محصولات با موفقیت تکمیل شد.", 'Success!', 20);
}

// --- توابع به‌روزرسانی (Update Functions) ---

/**
 * به‌روزرسانی قیمت و موجودی محصولات در WooCommerce بر اساس تغییرات شیت.
 */
function updateProductPricesCaller() {
  const [consumerKey, consumerSecret, siteUrl] = constants();
  // مطمئن شوید که شیت فعال (جاری) همان شیت Products باشد
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('خطا', 'شیت "Products" یافت نشد.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  updateProductPrices(consumerKey, consumerSecret, siteUrl, sheet);
}

/**
 * به‌روزرسانی محصولات/ورییشن‌ها در WooCommerce به صورت دسته‌ای.
 * @param {string} consumerKey کلید
 * @param {string} consumerSecret رمز
 * @param {string} siteUrl آدرس سایت
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet شیت Products
 */
function updateProductPrices(consumerKey, consumerSecret, siteUrl, sheet) {
  const dataRange = sheet.getDataRange();
  // بررسی کنید که آیا داده‌ای برای خواندن وجود دارد
  if (dataRange.getNumRows() <= 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast("هیچ محصولی برای به‌روزرسانی وجود ندارد.", 'توجه', 10);
    return;
  }

  const data = dataRange.getValues();
  const headers = data[0];

  // استخراج شاخص ستون‌ها بر اساس نام هدر
  const colIndex = Object.fromEntries(HEADERS.map(h => [h, headers.indexOf(h)]));

  // بررسی وجود ستون‌های حیاتی
  if (colIndex['Update'] === -1 || colIndex['ID'] === -1) {
    SpreadsheetApp.getUi().alert('خطا', 'ستون‌های حیاتی (Update یا ID) در شیت Products یافت نشد.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let simpleBatch = [];
  let variationBatches = {}; // Object: { parentId: [variationData] }
  const batchSize = 50; // استفاده از Batch Size کوچکتر برای اطمینان بیشتر
  let updatedCheckboxes = []; // آرایه برای نگهداری سلول‌های Update برای علامت‌گذاری (False)

  SpreadsheetApp.getActiveSpreadsheet().toast('شروع به‌روزرسانی موجودی و قیمت...', 'عملیات در حال انجام', -1);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const update = row[colIndex['Update']];

    if (update === true) {
      const productId = row[colIndex['ID']];
      const productType = row[colIndex['Type']];
      const regularPrice = row[colIndex['قیمت عادی']];
      const salePrice = row[colIndex['قیمت حراجی']];
      const stockQuantity = row[colIndex['موجودی']];
      const inStock = row[colIndex['موجود است؟']];
      const parentInfo = row[colIndex['Parent']]; // مثال: "id:123"

      // ساختار داده به‌روزرسانی
      const productData = {
        id: productId,
        // مقادیر خالی را به null/undefined تبدیل کنید تا API در صورت عدم تغییر، آن را نادیده بگیرد
        regular_price: regularPrice ? regularPrice.toString() : "",
        sale_price: salePrice ? salePrice.toString() : "",
        // موجودی: 1/0 برای موجود است؟
        stock_quantity: stockQuantity && !isNaN(parseInt(stockQuantity)) ? parseInt(stockQuantity) : null,
        manage_stock: true, // فرض کنید مدیریت موجودی فعال است
        stock_status: inStock == 1 ? 'instock' : 'outofstock'
      };

      if (productType === 'simple') {
        simpleBatch.push(productData);

        if (simpleBatch.length >= batchSize) {
          sendBatchRequest({ update: simpleBatch }, `${siteUrl}/wp-json/wc/v3/products/batch`, consumerKey, consumerSecret);
          simpleBatch = [];
        }
      } else if (productType === 'variation') {
        // استخراج Parent ID از ستون Parent (مثال: "id:123" -> 123)
        const parentMatch = parentInfo.toString().match(/id:(\d+)/);
        const parentId = parentMatch ? parentMatch[1] : null;

        if (parentId) {
          if (!variationBatches[parentId]) {
            variationBatches[parentId] = [];
          }

          variationBatches[parentId].push(productData);

          if (variationBatches[parentId].length >= batchSize) {
            sendBatchRequest({ update: variationBatches[parentId] }, `${siteUrl}/wp-json/wc/v3/products/${parentId}/variations/batch`, consumerKey, consumerSecret);
            variationBatches[parentId] = [];
          }
        } else {
          Logger.log(`Skipping variation ${productId}: Parent ID not found in Parent column.`);
        }
      }

      // جمع‌آوری موقعیت سلول برای بازنشانی تیک
      updatedCheckboxes.push({ row: i + 1, col: colIndex['Update'] + 1 });
    }
  }

  // ارسال باقی مانده Batch‌ها
  // --- بخش به‌روزرسانی نهایی در updateProductPrices ---

// ارسال باقی مانده Batch‌ها (مطابق با کد بهینه قبلی)
  if (simpleBatch.length > 0) {
    sendBatchRequest({ update: simpleBatch }, `${siteUrl}/wp-json/wc/v3/products/batch`, consumerKey, consumerSecret);
  }

  for (const parentId in variationBatches) {
    if (variationBatches[parentId].length > 0) {
      sendBatchRequest({ update: variationBatches[parentId] }, `${siteUrl}/wp-json/wc/v3/products/${parentId}/variations/batch`, consumerKey, consumerSecret);
    }
  }

// به‌روزرسانی تیک‌های Update به صورت **تک‌به‌تک** (تنها راه برای سلول‌های غیر متوالی)
// این روش بهینه تر از نوشتن خط به خط در حلقه اصلی است، زیرا فقط یک بار پس از اتمام API اجرا می شود.
  if (updatedCheckboxes.length > 0) {
    const checkValue = false;
    updatedCheckboxes.forEach(item => {
      sheet.getRange(item.row, item.col).setValue(checkValue);
    });
    // اطمینان از اعمال تغییرات به سرعت
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("همگام‌سازی قیمت و موجودی تکمیل شد.", 'Success!', 20);
}


/**
 * ارسال یک درخواست Batch به WooCommerce API.
 * @param {object} payload داده‌های Batch
 * @param {string} url آدرس API
 * @param {string} consumerKey کلید
 * @param {string} consumerSecret رمز
 */
function sendBatchRequest(payload, url, consumerKey, consumerSecret) {
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(consumerKey + ":" + consumerSecret)
    },
    muteHttpExceptions: true // برای مدیریت بهتر خطاها
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode >= 400) {
      Logger.log(`Batch Update Error ${responseCode}: ${response.getContentText()}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`خطای به‌روزرسانی دسته‌ای. کد: ${responseCode}`, 'خطا', 20);
    } else {
      Logger.log("Batch update successful.");
    }
  } catch (error) {
    Logger.log(`Critical Error in batch update: ${error.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`خطای بحرانی در به‌روزرسانی.`, 'خطا', 20);
  }
}

// --- توابع رابط کاربری (UI Functions) ---

/**
 * افزودن یک منوی سفارشی به Google Sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WordPress / WooCommerce')
      .addItem('1. واکشی محصولات (Fetch Products)', 'fetchProducts')
      .addSeparator()
      .addItem('2. همگام‌سازی قیمت و موجودی (Sync Prices & Stock)', 'updateProductPricesCaller')
      .addSeparator()
      .addItem('3. به‌روزرسانی تیک‌های Update (Reset Checkboxes)', 'refreshUpdateColumnCaller')
      .addToUi();
}

/**
 * Caller برای refreshUpdateColumn.
 */
function refreshUpdateColumnCaller() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  refreshUpdateColumn(sheet);
}

/**
 * بازسازی ستون "Update" با چک‌باکس‌ها.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet شیت مورد نظر
 */
function refreshUpdateColumn(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  const headers = headerRange.getValues()[0];
  const updateColIndex = headers.indexOf("Update") + 1;
  const updateColName = "Update";

  let targetColumn;

  if (updateColIndex > 0) {
    // ستون Update وجود دارد.
    targetColumn = updateColIndex;
  } else {
    // ستون Update وجود ندارد، آن را اضافه کنید.
    sheet.insertColumnAfter(lastColumn);
    targetColumn = lastColumn + 1;
    sheet.getRange(1, targetColumn).setValue(updateColName).setFontWeight('bold');
  }

  // اعمال چک‌باکس‌ها
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, targetColumn, lastRow - 1, 1);
    dataRange.clearContent();
    dataRange.insertCheckboxes();
  }

  sheet.autoResizeColumn(targetColumn);
  SpreadsheetApp.getActiveSpreadsheet().toast("ستون Update بازسازی شد.", 'Success!', 5);
}

/**
 * تنظیم "Update" به True هنگامی که ستون قیمت یا موجودی ویرایش می‌شود.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e رویداد ویرایش
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  const row = range.getRow();
  const col = range.getColumn();

  // فقط روی شیت Products و بعد از ردیف هدر
  if (sheet.getName() !== 'Products' || row === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updateCol = headers.indexOf("Update") + 1;

  // شاخص‌های ستون برای 'موجود است؟' (5), 'موجودی' (6), 'قیمت حراجی' (7), 'قیمت عادی' (8)
  const priceStockColumns = [5, 6, 7, 8];

  if (updateCol > 0 && priceStockColumns.includes(col)) {
    // جلوگیری از خطا در صورتی که سلول Update چک باکس نباشد
    const checkboxCell = sheet.getRange(row, updateCol);
    if (checkboxCell.getDataSourceUrl()) { // بررسی می کند که آیا یک چک باکس است
      checkboxCell.setValue(true);
    }
  }
}

