Of course. Here is the optimized README based on the code changes and enhancements, written in English.

-----

# WooCommerce-GoogleSheet-Sync (Optimized Version)

This project provides an advanced, high-performance Google Apps Script (GAS) designed to tightly integrate your **WooCommerce** store with **Google Sheets**. It allows you to reliably fetch product data, manage stock, update prices, and handle product variations directly from a spreadsheet.

## ✨ Key Features & Optimizations

This version is optimized for speed, reliability, and adherence to both Google Apps Script and WooCommerce API limits:

  * **⚡ High-Speed Fetching (Batch Writing):** Achieves significant performance gains by using **Batch Writing** to Google Sheets, updating hundreds of rows in a single operation instead of slow row-by-row updates.
  * **🔁 Continuation Management:** Includes built-in logic (`MAX_PRODUCTS_PER_RUN`) to prevent the script from hitting the Google Apps Script execution time limit (around 6 minutes). It saves the page number and resumes fetching automatically upon the next execution.
  * **📦 Efficient Batch API Updates:** Uses the WooCommerce REST API batch endpoints to update prices and stock for products and variations in small, controlled batches (50-100 per request) to respect API Rate Limits.
  * **🛡️ Robust Architecture:** Eliminates the inefficient, loop-based `findParentProductId` function, relying instead on the `Parent` column data for quick identification of variable products during updates.
  * **🛒 Full Variable Product Support:** Correctly fetches and updates variable products and their corresponding variations.
  * **⚙️ Custom UI:** Adds a dedicated **"WordPress / WooCommerce"** menu to the Google Sheet for easy access to all functions.

## 📋 Prerequisites

1.  A Google account.
2.  An active WooCommerce store.
3.  WooCommerce API enabled with a valid **Consumer Key** and **Consumer Secret** (Read/Write permissions required).

## 🛠️ Setup and Installation

### 1\. Enable WooCommerce REST API

1.  In your WordPress dashboard, go to **WooCommerce** \> **Settings** \> **Advanced** \> **REST API**.
2.  Click **Add Key**.
3.  Fill in the description, select the appropriate user, and set **Permissions** to **Read/Write**.
4.  Click **Generate API Key** and copy the resulting **Consumer Key** and **Consumer Secret**.

### 2\. Create and Populate the Google Sheet

1.  Create a new Google Sheet in your Google Drive.
2.  In the Google Sheet, go to **Extensions** \> **Apps Script**.
3.  Copy the entire optimized code and paste it into the Script editor, replacing any existing code.

### 3\. Configure API Credentials

Locate the `constants()` function at the top of the script and replace the placeholder values with your actual credentials:

```javascript
function constants() {
    const consumerKey = 'YOUR_CONSUMER_KEY';
    const consumerSecret = 'YOUR_CONSUMER_SECRET';
    // IMPORTANT: Site URL must NOT end with a trailing slash (e.g., 'https://yourstore.com')
    const siteUrl = 'https://YOUR_STORE_URL'; 
    return [consumerKey, consumerSecret, siteUrl];
}
```

### 4\. Save and Authorize

1.  Save the script (e.g., as "WooCommerceIntegration").
2.  In the Apps Script editor, select the **`onOpen`** function from the function dropdown and click the **Run (►)** button.
3.  Review and **authorize** the script to access your Google Sheet and external services (WooCommerce API).

## 🚀 Usage Guide

A custom menu named **"WordPress / WooCommerce"** will appear in your Google Sheet after successful authorization.

### 1\. Fetching Products

  * Click on **"WordPress / WooCommerce"** \> **"1. Fetch Products"**.
  * The script will create a new sheet named **"Products"** (or clear the existing one) and populate it with data from your WooCommerce store, including variations.
  * Columns like `ID`, `Type`, and `Parent` are automatically hidden for readability.
  * If your store has a large number of products, you may need to run this command multiple times. The script will remember the page where it left off and continue fetching.

### 2\. Syncing Prices & Stock

1.  In the **"Products"** sheet, modify the **"Sale price"**, **"Regular price"**, **"موجود است؟" (In Stock)**, or **"موجودی" (Stock Quantity)** columns.
2.  The **"Update"** checkbox column will automatically be checked for any row that is modified (via the `onEdit` trigger).
3.  Click on **"WordPress / WooCommerce"** \> **"2. Sync Prices & Stock"**.
4.  The script will read only the rows marked with **"Update = true"** and send the data in batches to your WooCommerce store.
5.  Upon successful synchronization, the **"Update"** checkboxes will automatically be set back to `false`.

### 3\. Refreshing Checkboxes

  * If the checkbox column is accidentally deleted or corrupted, you can restore it by clicking **"WordPress / WooCommerce"** \> **"3. Refresh Update Checkboxes"**.

## 🗂️ Key Script Functions

| Function | Description |
| :--- | :--- |
| `fetchProducts()` | Main execution for fetching products, managing pagination, and performing high-speed batch writing to the sheet. |
| `updateProductPrices()` | Reads the marked data and handles sending multiple batch update requests to the WooCommerce API. |
| `sendBatchRequest()` | Handles the authenticated POST request for WooCommerce batch updates, including basic error logging. |
| `onEdit(e)` | Automatically checks the `Update` column checkbox when a price or stock column is modified. |
| `onOpen()` | Creates the custom "WordPress / WooCommerce" menu in the sheet UI. |

## ⚠️ Notes

  * **API Throttling:** The script incorporates brief delays and manages updates in batches to minimize the risk of hitting WooCommerce API Rate Limits.
  * **Data Integrity:** Ensure that the API keys have the necessary **Read/Write** permissions to avoid update failures.
  * **Troubleshooting:** The script includes logging for errors (`Logger.log()`). If issues occur, check the **Executions** log in the Apps Script editor.
