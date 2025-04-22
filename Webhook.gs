/**
 * Mystery Order Email Automation - Webhook Components
 * Functions for handling email tracking and API endpoints
 * @version 1.0
 * @lastModified 2025-04-18
 */

/**
 * Web app doGet function that handles tracking pixel requests
 * This function will be called when the tracking pixel is loaded
 */
function doGet(e) {
  try {
    // Get the tracking ID from the request
    const trackingId = e.parameter.id;
    const action = e.parameter.action;
    
    if (action === 'open' && trackingId) {
      // Record that the email was opened
      recordEmailOpen(trackingId);
    }
    
    // Return a 1x1 transparent GIF
    return ContentService.createTextOutput('GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;')
      .setMimeType(ContentService.MimeType.IMAGE);
  } catch (error) {
    Logger.log(`Error processing tracking request: ${error.message}`);
    return ContentService.createTextOutput('Error').setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Record that an email was opened
 * @param {string} trackingId The tracking ID from the pixel request
 */
function recordEmailOpen(trackingId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    
    if (!trackingSheet) {
      Logger.log('Tracking sheet not found');
      return;
    }
    
    // Get all tracking data
    const data = trackingSheet.getDataRange().getValues();
    
    // Skip the header row and find the tracking ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === trackingId) {
        // Only update if this is the first time the email was opened
        if (data[i][5] === 'No') {
          const now = new Date();
          
          // Update the row to indicate the email was opened
          trackingSheet.getRange(i + 1, 5).setValue(now); // Open Date
          trackingSheet.getRange(i + 1, 6).setValue('Yes'); // Opened
        } else {
          // Increment the views counter
          let currentViews = data[i][6] || 0;
          trackingSheet.getRange(i + 1, 7).setValue(currentViews + 1);
        }
        break;
      }
    }
  } catch (error) {
    Logger.log(`Error recording email open: ${error.message}`);
  }
}
