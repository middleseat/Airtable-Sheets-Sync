# Google Sheets to Airtable Sync

A tool to synchronize donation data between Google Sheets and Airtable.

## Overview

This project contains Google Apps Script code that:

1. Fetches ActBlue donation form URLs from Airtable
2. Extracts form slugs from these URLs
3. Matches them with form names in Google Sheets
4. Sums donation amounts for each form
5. Updates the corresponding "Raised" field in Airtable

## Features

- **Automatic Syncing**: Updates automatically when the Google Sheet is edited (rate-limited to once per hour)
- **Manual Syncing**: Can be triggered manually from a custom menu in Google Sheets (not rate-limited)
- **Error Handling**: Comprehensive error logging and display
- **Rate Limiting**: Prevents excessive API calls to Airtable
- **Data Validation**: Ensures only valid data is processed and updated

## Files

- [SheetsToAirtableSync.js](src/google-apps-script/SheetsToAirtableSync.js): The main Google Apps Script code
- [DEPLOYMENT.md](src/google-apps-script/DEPLOYMENT.md): Instructions for deploying the script
- [CONFIGURATION.md](CONFIGURATION.md): Configuration guide for setting up the sync

## Getting Started

1. Follow the setup instructions in [CONFIGURATION.md](CONFIGURATION.md)
2. Deploy the script using [DEPLOYMENT.md](src/google-apps-script/DEPLOYMENT.md)
3. Run a manual sync to test the integration

## How It Works

The script connects to Airtable's API to fetch and update records. It uses the ActBlue form slugs to match records between the two systems, and calculates the total donations for each form from the Google Sheet data.

Data is automatically synced when edits are made to the Google Sheet, but to prevent API abuse, these automatic syncs are limited to once per hour. Manual syncs can be triggered at any time through the custom menu. 