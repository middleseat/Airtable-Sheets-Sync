/**
 * Google Sheets to Airtable Sync
 * This script syncs data between Google Sheets and Airtable when:
 * 1. The user edits the sheet (limited to once per hour)
 * 2. The user manually triggers it from the menu (no limit)
 */

// Configuration constants
const AIRTABLE_API_KEY = "patmOcjxRihjKvEsJ.56abda1b858edd6056d5d000ad0fd68ebcd3a36a085d5fda7e1652b6b47456c2"; // Airtable API key
const AIRTABLE_BASE_ID = "appGDzLGgrYmcW9R1"; // The base ID for your Airtable
const AIRTABLE_TABLE_ID = "tblTo3zgLYNby9UnL"; // The 2024 P2P Texting table ID
const SHEET_NAME = "raw_import"; // The sheet name containing the donor data
const ACTBLUE_URL_PREFIX = "https://secure.actblue.com/donate/"; // Prefix to remove from ActBlue URLs

// Rate limiting constants
const RATE_LIMIT_HOURS = 1; // Only auto-sync once per hour
const RATE_LIMIT_PROP_KEY = "lastAutoSyncTime"; // For storing the last sync time

/**
 * Creates a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Airtable Sync')
    .addItem('Sync Data Now', 'manualSyncData')
    .addItem('View Sync Logs', 'viewLogs')
    .addToUi();
}

/**
 * Triggered when the sheet is edited.
 * Limits syncing to once per hour for automatic updates.
 */
function onEdit(e) {
  // Only trigger if the edit happens in the raw_import sheet
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) {
    return;
  }
  
  // Check if we should respect the rate limit
  if (shouldRespectRateLimit()) {
    logMessage("Edit detected, but skipping auto-sync due to rate limiting");
    return;
  }
  
  // Update the last sync time
  updateLastSyncTime();
  
  // Run the sync
  logMessage("Edit detected, running auto-sync");
  syncData();
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
    logMessage("Starting sync process...");
    
    // Fetch ActBlue URLs from Airtable
    const airtableRecords = fetchAirtableRecords();
    if (!airtableRecords || airtableRecords.length === 0) {
      logMessage("No records found in Airtable, sync complete");
      return;
    }
    
    // Get Google Sheet data - pass airtableRecords to filter by form slugs
    const sheetData = getSheetData(airtableRecords);
    if (!sheetData || sheetData.length === 0) {
      logMessage("No matching data found in Google Sheet, sync complete");
      return;
    }
    
    // Process and aggregate the data
    const processedData = processData(airtableRecords, sheetData);
    if (!processedData || processedData.length === 0) {
      logMessage("No matches found between Airtable and Sheet data, sync complete");
      return;
    }
    
    // Update Airtable with the processed data
    updateAirtable(processedData);
    
    logMessage("Sync completed successfully");
  } catch (error) {
    logError("Error in sync process: " + error.message);
  }
}

/**
 * Fetches records from Airtable API.
 * @return {Array} Array of Airtable records with their IDs and ActBlue URLs
 */
function fetchAirtableRecords() {
  logMessage("Fetching records from Airtable...");
  
  try {
    // Prepare request options
    const url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${AIRTABLE_TABLE_ID}`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${AIRTABLE_API_KEY}`,
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
      const actBlueUrl = record.fields["ActBlue Page"] || 
                         record.fields[`fldxiPzVeVCUXsnq1`] || '';
      
      // Extract form slug by removing the ActBlue URL prefix
      let formSlug = '';
      if (actBlueUrl && actBlueUrl.indexOf(ACTBLUE_URL_PREFIX) !== -1) {
        formSlug = actBlueUrl.replace(ACTBLUE_URL_PREFIX, '');
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
 * @param {Array} processedData - Data to update in Airtable
 */
function updateAirtable(processedData) {
  logMessage(`Updating ${processedData.length} Airtable records...`);
  
  let successCount = 0;
  let errorCount = 0;
  
  try {
    // Process each record
    processedData.forEach((record, index) => {
      try {
        // Add slight delay between requests to avoid rate limiting
        if (index > 0) {
          Utilities.sleep(200);
        }
        
        // Prepare request options
        const url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${AIRTABLE_TABLE_ID}/${record.id}`;
        const payload = {
          fields: {
            // Use both field names and IDs for robustness
            "Raised": record.raised,
            "fldoNVriyq6Z0YEE7": record.raised,
            "Donations": record.donations,
            "fldgEEtQMx96RUqds": record.donations
          }
        };
        
        const options = {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${AIRTABLE_API_KEY}`,
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
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create a new edit trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  logMessage("Edit trigger has been set up");
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
 * Displays logs in a modal dialog.
 */
function viewLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName("Logs");
  
  if (!logsSheet) {
    SpreadsheetApp.getUi().alert("No logs found. Run a sync first to generate logs.");
    return;
  }
  
  // Get log data
  const dataRange = logsSheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length <= 1) {
    SpreadsheetApp.getUi().alert("No log entries found.");
    return;
  }
  
  // Format the logs for display (show the most recent 50 entries)
  const maxEntries = Math.min(50, values.length - 1);
  const startIndex = Math.max(1, values.length - maxEntries);
  
  let logText = "<html><body><table style='border-collapse: collapse; width: 100%;'>";
  logText += "<tr style='background-color: #f2f2f2;'><th style='border: 1px solid #ddd; padding: 8px; text-align: left;'>Timestamp</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left;'>Type</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left;'>Message</th></tr>";
  
  for (let i = startIndex; i < values.length; i++) {
    const row = values[i];
    const timestamp = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const type = row[1];
    const message = row[2];
    
    const rowStyle = type === "ERROR" ? "color: red;" : "";
    logText += `<tr style='${rowStyle}'><td style='border: 1px solid #ddd; padding: 8px;'>${timestamp}</td><td style='border: 1px solid #ddd; padding: 8px;'>${type}</td><td style='border: 1px solid #ddd; padding: 8px;'>${message}</td></tr>`;
  }
  
  logText += "</table></body></html>";
  
  // Display logs in a modal dialog
  const htmlOutput = HtmlService
    .createHtmlOutput(logText)
    .setWidth(800)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Sync Logs (Recent " + maxEntries + " entries)");
} 