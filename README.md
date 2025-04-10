# Google Sheets to Airtable Sync

A tool for one-way synchronization of donation data from Google Sheets to Airtable. This script pulls data from your Google Sheet and updates your Airtable records, not vice versa.

## Setup Instructions

### 1. Create a new Google Apps Script project

1. Open your Google Sheet
2. Select Extensions > Apps Script
3. Copy the contents of [SheetsToAirtableSync.js](SheetsToAirtableSync.js) into the script editor
4. Save the project

> **Note:** Formal deployment is not needed. The script automatically works with your sheet once set up. If you share your sheet with others, users with edit access will also be able to view and edit the script.

### 2. Configure the script

Edit the configuration section at the top of the script:

```javascript
// Configuration
const AIRTABLE_BASE_ID = "your_base_id"; // The base ID for your Airtable
const AIRTABLE_TABLE_ID = "your_table_id"; // The table ID
const SHEET_NAME = "raw_import"; // The sheet name containing the donor data
const RATE_LIMIT_HOURS = 0.25; // Only auto-sync once per 15 minutes

// Airtable field names
const AIRTABLE_ACTBLUE_PAGE_FIELD = "ActBlue Page"; // Field name for ActBlue page URL
const AIRTABLE_RAISED_FIELD = "Raised"; // Field name for amount raised
const AIRTABLE_DONATIONS_FIELD = "Donations"; // Field name for number of donations
```

### 3. Set up your Airtable API key

1. Get your Airtable API key from your Airtable account
2. In the Apps Script editor, go to Project Settings (gear icon)
3. Under "Script Properties", click "Add script property"
4. Add a property with:
   - Property Name: `AIRTABLE_API_KEY`
   - Value: Your Airtable API key
5. Click "Save script properties"

### 4. Enable the script

1. Refresh your Google Sheet
2. A new menu item "Airtable Sync" will appear
3. Click "Airtable Sync" > "Install Auto-Sync"
4. When prompted, authorize the script to access your data
5. You'll see a confirmation message when auto-sync is installed

## Using the Script

The script adds a custom menu to your Google Sheet with the following options:

- **Sync Data Now**: Manually trigger a sync (bypasses rate limiting)
- **Install Auto-Sync**: Set up the trigger for automatic syncing (one-time setup)
- **View Logs**: Opens the Logs sheet to view sync history

## How It Works

The script connects to Airtable's API to fetch and update records. It uses the ActBlue form slugs to match records between the two systems, and calculates the total donations and donation counts for each form from the Google Sheet data.

Data is automatically synced when edits are made to the Google Sheet, but these automatic syncs are limited to once every 15 minutes. Manual syncs can be triggered at any time through the custom menu. 