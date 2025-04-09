# Configuration Guide

## Google Sheets Setup

1. Ensure your Google Sheet has a sheet named `raw_import` with the following columns:
   - `form_name` (column D) - This should match the ActBlue form slug
   - `dollars_raised` (column F) - The donation amount
   - `num_of_donations` (column E) - The number of donations

## Airtable Setup

1. Get your Airtable API key from your Airtable account settings
2. Note your base ID and table ID:
   - Base ID: `appGDzLGgrYmcW9R1` (or your custom base ID)
   - Table ID: `tblTo3zgLYNby9UnL` (or your custom table ID)
3. Make sure your Airtable table has the following fields:
   - "ActBlue Page" (field ID: `fldxiPzVeVCUXsnq1`) - URLs in format `https://secure.actblue.com/donate/form-slug`
   - "Raised" (field ID: `fldoNVriyq6Z0YEE7`) - This field will be updated with the total amount raised
   - "Donations" (field ID: `fldgEEtQMx96RUqds`) - This field will be updated with the total number of donations

## Script Configuration

1. In the `SheetsToAirtableSync.js` file, update the following constants if needed:
   ```javascript
   const AIRTABLE_API_KEY = "YOUR_AIRTABLE_API_KEY"; // Replace with your actual key (already updated in the script)
   const AIRTABLE_BASE_ID = "appGDzLGgrYmcW9R1"; // Update if using a different base
   const AIRTABLE_TABLE_ID = "tblTo3zgLYNby9UnL"; // Update if using a different table
   const SHEET_NAME = "raw_import"; // The sheet containing donation data
   const ACTBLUE_URL_PREFIX = "https://secure.actblue.com/donate/"; // The prefix of ActBlue URLs
   const RATE_LIMIT_HOURS = 1; // How often auto-sync can occur (in hours)
   ```

2. Ensure you have the necessary permissions:
   - Edit access to the Google Sheet
   - API access to the Airtable base

## Understanding the Data Flow

1. The script fetches ActBlue URLs from Airtable and extracts form slugs
2. It then gets donation data from the Google Sheet
3. For each form_name in the sheet that matches an ActBlue form slug:
   - It sums the dollars_raised values
   - It sums the num_of_donations values
4. These totals are then updated in the Airtable "Raised" and "Donations" fields for the matching records

## Customization Options

- To change the rate limit for automatic syncs, modify the `RATE_LIMIT_HOURS` constant
- To track different fields, update the column indices in the `getSheetData()` function
- To update different Airtable fields, modify the payload in the `updateAirtable()` function 