# WooCommerce-GoogleSheet-Sync
This project provides a Google Apps Script to integrate your WooCommerce store with Google Sheets. It allows you to fetch product data, update prices, manage stock, and manage your WooCommerce products directly from a Google Sheet.

## Features

*   **Fetch Products:** Retrieves product data from your WooCommerce store and populates a Google Sheet.
*   **Update Prices:** You can update product prices in WooCommerce by modifying the Google Sheet.
*   **Variable Product Support:** Handles variable products and their variations.
*   **Custom Menu:** Adds a custom menu to Google Sheets for easy access to the script's functions.
*   **Update Checkboxes:** Adds checkboxes to select products for price updates easily.
*   **Batch Processing:** Updates product prices in batches to avoid exceeding API limits.

## Prerequisites

*   A Google account.
*   A WooCommerce store.
*   WooCommerce API enabled with valid Consumer Key and Consumer Secret.

## Setup

### 1. Create a Google Sheet

Create a new Google Sheet in your Google Drive.

### 2. Open Script Editor

In the Google Sheets, go to "Tools" > "Script editor".

### 3. Copy and Paste the Code

Copy the entire code from `WooCommerce-GoogleSheets-Sync.js` and paste it into the Script editor.

### 4. Configure WooCommerce API Credentials

Modify the `constants()` function in the script to include your WooCommerce API credentials:

```javascript
// WooCommerce-GoogleSheets-Sync.js
function constants() {
    var consumerKey = 'YOUR_CONSUMER_KEY';
    var consumerSecret = 'YOUR_CONSUMER_SECRET';
    var siteUrl = 'YOUR_SITE_URL'; // e.g., 'https://yourstore.com/'
    return [consumerKey, consumerSecret, siteUrl];
  }
```
Replace YOUR_CONSUMER_KEY, YOUR_CONSUMER_SECRET, and YOUR_SITE_URL with your actual WooCommerce API credentials and store URL.

5. Save the Script
Save the script with a descriptive name (e.g., "WooCommerceIntegration").

6. Run the onOpen() Function
Select the onOpen function in the Script editor's function dropdown.
Click the "Run" button (play icon).
Authorize the script to access your Google Sheet.
7. Custom Menu
A custom menu named "WordPress" will be added to the Google Sheet. It contains the following options:

Fetch Products: Fetches product data from WooCommerce and populates the sheet.
Send Updated Prices: Updates product prices in WooCommerce based on the changes made in the sheet.
Usage
Fetching Products
Click on "WordPress" > "Fetch Products" in the Google Sheet menu.
The script will create a new sheet named "Products" (or clear the existing one) and populate it with product data from your WooCommerce store.
A temporary sheet named "TempFetch" is created to keep track of the page number during fetching. This sheet is automatically deleted after the process is complete.
Updating Prices
In the "Products" sheet, modify the "Sale price" and/or "Regular price" columns for the products you want to update.
Check the "Update" column for each product you've modified.
Click on "WordPress" > "Send Updated Prices" in the Google Sheet menu.
The script will update the prices in your WooCommerce store based on the changes you made in the sheet.
The "Update" checkboxes will be automatically unchecked after the update.
Functions
constants(): Configures WooCommerce API access information.
fetchProducts(): Fetches products from WooCommerce and populates the Google Sheet.
fetchProductBatch(siteUrl, consumerKey, consumerSecret, page, perPage): Retrieves a batch of products from the WooCommerce API.
fetchVariations(siteUrl, consumerKey, consumerSecret, productId): Retrieves variations for a variable product.
getVariationAttributes(variation): Combines variation attributes into a string.
makeAuthenticatedRequest(url, params, consumerKey, consumerSecret): Sends an authenticated request to the WooCommerce API.
hideColumns(sheet, columnsToHide, headers): Hides specified columns in the Google Sheet.
refreshUpdateColumn(): Creates or refreshes the "Update" checkbox column.
onOpen(): Creates a custom menu in Google Sheets.
onEdit(e): Automatically checks the "Update" checkbox when a price is modified.
updateProductPrices(consumerKey, consumerSecret, siteUrl): Updates product prices in WooCommerce based on the Google Sheet.
sendBatchRequest(batch, url, consumerKey, consumerSecret): Sends a batch request to the WooCommerce API to update product data.
findParentProductId(data, variationId, idCol, typeCol): Finds the parent product ID for a given variation ID.
Notes
The script uses batch processing to update product prices, which helps to avoid exceeding API limits. The batch size is set to 100 products per batch.
The script automatically resizes columns and hides unnecessary columns for better readability.
The TempFetch sheet is used to store the current page number during the fetch process. It is automatically created and deleted by the script.
The onEdit(e) function automatically checks the "Update" checkbox when a price is modified in the "Sale price" or "Regular price" columns.
Ensure that your WooCommerce API keys have the necessary permissions to read and write product data.
Error Handling
The script includes basic error handling to log errors to the Logger. Check the script editor's execution log for any errors that may occur during the process.

Disclaimer
This script is provided as-is and without any warranty. Use it at your own risk. The author is not responsible for any data loss or other issues that may arise from using this script.
