/**
 * Google Sheets to Airtable ActBlueSync
 * Version: 1.1.0
 * Author(s): Ryan Mioduski
 *
 * Important:
 * This script syncs data between Google Sheets and Airtable, specifically for tracking
 * ActBlue donation data. It automatically updates Airtable records with the total amount
 * raised and number of donations for each ActBlue form.
 *
 * Before using the script, you need to:
 * 1. Set up your Airtable API key in the script properties
 * 2. Configure the correct Airtable base ID and table ID
 * 3. Ensure your Google Sheet has the correct sheet name and column structure
 */

// Configuration
// Airtable destinations (add IDs to the second object when needed)
const AIRTABLE_TARGETS = [
  { baseId: "appGDzLGgrYmcW9R1", tableId: "tblTo3zgLYNby9UnL" }, // Primary destination
  { baseId: "", tableId: "" } // Secondary destination – leave blank until ready
];
const SHEET_NAME = "raw_import"; // The sheet name containing the donor data
const RATE_LIMIT_HOURS = 0.25; // Only auto-sync once per 15 minutes

// Airtable field names
const AIRTABLE_ACTBLUE_PAGE_FIELD = "ActBlue Page"; // Field name for ActBlue page URL
const AIRTABLE_RAISED_FIELD = "Raised"; // Field name for amount raised
const AIRTABLE_DONATIONS_FIELD = "Donations"; // Field name for number of donations
// End Configuration

// Internal constants
const RATE_LIMIT_PROP_KEY = "lastAutoSyncTime"; // For storing the last sync time
const AIRTABLE_API_KEY_PROP = "AIRTABLE_API_KEY"; // Property key for storing the API key

/**
 * Creates a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Airtable Sync')
    .addItem('Sync Data Now', 'manualSyncData')
    .addItem('Install Hourly Auto-Sync', 'installAutoSync')
    .addSeparator()
    .addItem('View Logs', 'viewLogs')
    .addToUi();
}

/**
 * Manual sync function that bypasses rate limiting.
 * Called from the custom menu.
 */
function manualSyncData() {
  logMessage("Manual sync triggered by user");
  syncData();
}

/**
 * Creates an hourly time-based trigger for auto-sync.
 * This is a ONE-TIME SETUP that must be run to enable auto-sync.
 */
function installAutoSync() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'hourlyAutoSync') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new installable hourly trigger
  ScriptApp.newTrigger('hourlyAutoSync')
    .timeBased()
    .everyHours(1)
    .create();
  
  // Show confirmation
  const ui = SpreadsheetApp.getUi();
  ui.alert("Hourly Auto-Sync Installed", 
           "Auto-sync has been successfully installed! The spreadsheet will now automatically " +
           "sync with Airtable every hour, regardless of how the data changes.\n\n" +
           "P.S. You can safely run 'Install Hourly Auto-Sync' again if needed - it will update the existing trigger without creating duplicates.", 
           ui.ButtonSet.OK);
  
  logMessage("Hourly auto-sync trigger installed successfully");
}

/**
 * This function is called by the hourly time-based trigger.
 * It syncs data regardless of how it was updated (manually or by another script).
 */
function hourlyAutoSync() {
  logMessage("Hourly auto-sync triggered");
  
  // Check if we should respect the rate limit (keeping this as a safeguard)
  if (shouldRespectRateLimit()) {
    logMessage("Hourly auto-sync triggered, but skipping due to rate limiting");
    return;
  }
  
  // Update the last sync time
  updateLastSyncTime();
  
  // Run the sync
  syncData();
}

/**
 * Checks if we should respect the rate limit.
 * @return {boolean} True if we should respect the rate limit
 */
function shouldRespectRateLimit() {
  const props = PropertiesService.getScriptProperties();
  const lastSyncTimeStr = props.getProperty(RATE_LIMIT_PROP_KEY);
  
  if (!lastSyncTimeStr) {
    return false; // No last sync time stored, allow sync
  }
  
  const lastSyncTime = new Date(lastSyncTimeStr);
  const now = new Date();
  const hoursDiff = (now - lastSyncTime) / (1000 * 60 * 60);
  
  return hoursDiff < RATE_LIMIT_HOURS;
}

/**
 * Updates the last sync time to now.
 */
function updateLastSyncTime() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(RATE_LIMIT_PROP_KEY, new Date().toISOString());
}

/**
 * Main function to sync data from Google Sheets to Airtable.
 */
function syncData() {
  try {
    logMessage("Starting multi-destination sync process...");

    AIRTABLE_TARGETS.forEach(cfg => {
      // Skip any target that hasn't been configured yet
      if (!cfg.baseId || !cfg.tableId) {
        logMessage("Skipping target with blank Base/Table IDs");
        return;
      }

      // Step 1: Fetch ActBlue URLs from this Airtable table
      const airtableRecords = fetchAirtableRecords(cfg);
      if (!airtableRecords || airtableRecords.length === 0) {
        logMessage(`No records found in Airtable for base ${cfg.baseId}, skipping`);
        return;
      }

      // Step 2: Get matching Google Sheet rows
      const sheetData = getSheetData(airtableRecords);
      if (!sheetData || sheetData.length === 0) {
        logMessage("No matching data found in Google Sheet for this target, skipping");
        return;
      }

      // Step 3: Aggregate
      const processedData = processData(airtableRecords, sheetData);
      if (!processedData || processedData.length === 0) {
        logMessage("No matches found between Airtable and Sheet data for this target, skipping");
        return;
      }

      // Step 4: Update Airtable
      updateAirtable(cfg, processedData);
    });

    logMessage("Sync completed for all configured targets");
  } catch (error) {
    logError("Error in sync process: " + error.message);
  }
}

/**
 * Fetches records from Airtable API.
 * @param {Object} cfg - Configuration object for the Airtable target
 * @return {Array} Array of Airtable records with their IDs and ActBlue URLs
 */
function fetchAirtableRecords(cfg) {
  logMessage(`Fetching records from Airtable (base: ${cfg.baseId}, table: ${cfg.tableId})...`);
  
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty(AIRTABLE_API_KEY_PROP);
    
    if (!apiKey) {
      logError("Airtable API key not found in script properties");
      return [];
    }
    
    // Prepare request options
    const url = `https://api.airtable.com/v0/${cfg.baseId}/${cfg.tableId}`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    };
    
    // Make the request
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.records) {
      logError("No records found in Airtable response");
      return [];
    }
    
    // Extract and process the records
    const records = data.records.map(record => {
      const id = record.id;
      const actBlueUrl = record.fields[AIRTABLE_ACTBLUE_PAGE_FIELD] || '';
      
      // Extract form slug by removing the ActBlue URL prefix
      let formSlug = '';
      if (actBlueUrl && actBlueUrl.indexOf("https://secure.actblue.com/donate/") !== -1) {
        formSlug = actBlueUrl.replace("https://secure.actblue.com/donate/", '');
      }
      
      return {
        id,
        actBlueUrl,
        formSlug
      };
    }).filter(record => record.formSlug); // Only keep records with a valid form slug
    
    logMessage(`Found ${records.length} valid ActBlue form slugs in Airtable`);
    return records;
    
  } catch (error) {
    logError("Error fetching Airtable records: " + error.message);
    return [];
  }
}

/**
 * Gets donor data from the Google Sheet.
 * @param {Array} airtableRecords - Records from Airtable to filter form_names
 * @return {Array} Array of donation data from the sheet
 */
function getSheetData(airtableRecords) {
  logMessage("Retrieving data from Google Sheets...");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      logError(`Sheet '${SHEET_NAME}' not found`);
      return [];
    }
    
    // Extract form slugs from Airtable records for filtering
    const formSlugs = airtableRecords.map(record => record.formSlug);
    logMessage(`Will filter sheet data for ${formSlugs.length} form slugs`);
    
    // Get all data
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      logError("No data found in sheet or only header row present");
      return [];
    }
    
    // Extract headers
    const headers = values[0];
    
    // Find the indices for form_name, dollars_raised, and num_of_donations columns
    const formNameIndex = headers.indexOf('form_name');
    const dollarsRaisedIndex = headers.indexOf('dollars_raised');
    const numDonationsIndex = headers.indexOf('num_of_donations');
    
    if (formNameIndex === -1 || dollarsRaisedIndex === -1 || numDonationsIndex === -1) {
      logError("Could not find required columns in sheet headers");
      return [];
    }
    
    logMessage(`Processing ${values.length} rows from sheet, filtering for ${formSlugs.length} form slugs...`);
    
    // Transform data into objects with named properties, but only for matching form_names
    const sheetData = [];
    let processedRows = 0;
    let matchedRows = 0;
    
    for (let i = 1; i < values.length; i++) {
      processedRows++;
      const row = values[i];
      const formName = row[formNameIndex];
      
      // Only process rows with form_names that match our Airtable form slugs
      if (formName && formSlugs.includes(formName)) {
        matchedRows++;
        const dollarsRaised = parseFloat(row[dollarsRaisedIndex]) || 0;
        const numDonations = parseInt(row[numDonationsIndex]) || 0;
        
        if (dollarsRaised > 0 || numDonations > 0) {
          sheetData.push({
            formName,
            dollarsRaised,
            numDonations
          });
        }
      }
      
      // Log progress for large datasets
      if (processedRows % 50000 === 0) {
        logMessage(`Processed ${processedRows} rows, found ${matchedRows} matches so far...`);
      }
    }
    
    logMessage(`Found ${sheetData.length} matching donation records in Google Sheet (from ${values.length-1} total rows)`);
    return sheetData;
    
  } catch (error) {
    logError("Error retrieving sheet data: " + error.message);
    return [];
  }
}

/**
 * Processes and aggregates the data by matching ActBlue form slugs.
 * @param {Array} airtableRecords - Records fetched from Airtable
 * @param {Array} sheetData - Data from Google Sheets
 * @return {Array} Processed data ready for updating Airtable
 */
function processData(airtableRecords, sheetData) {
  logMessage("Processing and aggregating data...");
  
  try {
    // Create maps to store the sums for each form name
    const formTotals = {};
    const formDonations = {};
    
    // Aggregate dollars raised and number of donations by form name
    sheetData.forEach(item => {
      const formName = item.formName;
      const amount = item.dollarsRaised;
      const donations = item.numDonations;
      
      if (!formTotals[formName]) {
        formTotals[formName] = 0;
        formDonations[formName] = 0;
      }
      
      formTotals[formName] += amount;
      formDonations[formName] += donations;
    });
    
    // Match the aggregated data with Airtable records
    const processedData = [];
    
    airtableRecords.forEach(record => {
      const formSlug = record.formSlug;
      
      if (formTotals[formSlug] !== undefined || formDonations[formSlug] !== undefined) {
        processedData.push({
          id: record.id,
          raised: formTotals[formSlug] || 0,
          donations: formDonations[formSlug] || 0
        });
      }
    });
    
    logMessage(`Matched and processed ${processedData.length} records`);
    return processedData;
    
  } catch (error) {
    logError("Error processing data: " + error.message);
    return [];
  }
}

/**
 * Updates Airtable records with the processed data.
 * @param {Object} cfg - Configuration object for the Airtable target
 * @param {Array} processedData - Data to update in Airtable
 */
function updateAirtable(cfg, processedData) {
  logMessage(`Updating ${processedData.length} Airtable records in base ${cfg.baseId}...`);
  
  let successCount = 0;
  let errorCount = 0;
  
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty(AIRTABLE_API_KEY_PROP);
    
    if (!apiKey) {
      logError("Airtable API key not found in script properties");
      return;
    }
    
    // Process each record
    processedData.forEach((record, index) => {
      try {
        // Add slight delay between requests to avoid rate limiting
        if (index > 0) {
          Utilities.sleep(200);
        }
        
        // Prepare request options
        const url = `https://api.airtable.com/v0/${cfg.baseId}/${cfg.tableId}/${record.id}`;
        const payload = {
          fields: {
            [AIRTABLE_RAISED_FIELD]: record.raised,
            [AIRTABLE_DONATIONS_FIELD]: record.donations
          }
        };
        
        const options = {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(payload)
        };
        
        // Make the request
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          successCount++;
          logMessage(`Updated record ${record.id} with amount ${record.raised} and ${record.donations} donations`);
        } else {
          errorCount++;
          logError(`Failed to update record ${record.id}: Response code ${responseCode}`);
        }
      } catch (error) {
        errorCount++;
        logError(`Error updating record ${record.id}: ${error.message}`);
      }
    });
    
    logMessage(`Update complete: ${successCount} successful, ${errorCount} failed`);
    
  } catch (error) {
    logError("Error updating Airtable: " + error.message);
  }
}

/**
 * Sets up an edit trigger for the sheet.
 */
function setupTrigger() {
  // This function is now unused but kept for backward compatibility
  logMessage("Manual authorization is now handled automatically. Just run 'Sync Data Now' from the menu.");
}

/**
 * Logs a message to a dedicated "Logs" sheet.
 * @param {string} message - The message to log
 */
function logMessage(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName("Logs");
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet("Logs");
    logsSheet.appendRow(["Timestamp", "Type", "Message"]);
    logsSheet.setFrozenRows(1);
    
    // Format header row
    logsSheet.getRange("A1:C1").setFontWeight("bold");
    logsSheet.setColumnWidth(1, 200); // Timestamp
    logsSheet.setColumnWidth(2, 100); // Type
    logsSheet.setColumnWidth(3, 500); // Message
  }
  
  logsSheet.appendRow([new Date(), "INFO", message]);
}

/**
 * Logs an error message to the "Logs" sheet.
 * @param {string} errorMessage - The error message to log
 */
function logError(errorMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName("Logs");
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet("Logs");
    logsSheet.appendRow(["Timestamp", "Type", "Message"]);
    logsSheet.setFrozenRows(1);
    
    // Format header row
    logsSheet.getRange("A1:C1").setFontWeight("bold");
    logsSheet.setColumnWidth(1, 200); // Timestamp
    logsSheet.setColumnWidth(2, 100); // Type
    logsSheet.setColumnWidth(3, 500); // Message
  }
  
  // Add error row with red text
  logsSheet.appendRow([new Date(), "ERROR", errorMessage]);
  const lastRow = logsSheet.getLastRow();
  logsSheet.getRange(lastRow, 1, 1, 3).setFontColor("red");
}

/**
 * Displays logs by activating the Logs sheet tab.
 */
function viewLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName("Logs");
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet("Logs");
    logsSheet.appendRow(["Timestamp", "Type", "Message"]);
    logsSheet.setFrozenRows(1);
    
    // Format header row
    logsSheet.getRange("A1:C1").setFontWeight("bold");
    logsSheet.setColumnWidth(1, 200); // Timestamp
    logsSheet.setColumnWidth(2, 100); // Type
    logsSheet.setColumnWidth(3, 500); // Message
    
    SpreadsheetApp.getUi().alert("Created new Logs sheet.");
  }
  
  // Activate the Logs sheet to make it visible
  logsSheet.activate();
} 