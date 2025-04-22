/**
 * Mystery Order Email Automation - Email Components
 * Functions for email sending and tracking
 * @version 1.0
 * @lastModified 2025-04-18

 */

/**
 * Sends a notification email using the provided template and data
 * @param {Object} htmlTemplate The HTML template to use
 * @param {Object} data Data to populate the template
 * @param {String} subject Email subject
 * @return {Boolean} Whether the email was sent successfully
 */
function sendNotificationEmail(htmlTemplate, data, subject) {
  try {
    // Populate the template
    for (const key in data) {
      htmlTemplate[key] = data[key];
    }
    
    // Evaluate the template and get the HTML content
    let htmlForEmail = htmlTemplate.evaluate().getContent();
    
    // Add tracking pixel to the HTML content
    const trackingPixel = generateTrackingPixel(data.sellerEmail, data.sellerId);
    htmlForEmail += trackingPixel;
    
    // Handle multiple email addresses
    const emailAddresses = data.sellerEmail.split(',').map(email => email.trim());
    let successCount = 0;
    
    // Send email to each recipient
    for (const email of emailAddresses) {
      if (isValidEmail(email)) {
        GmailApp.sendEmail(
          email,
          subject,
          `Mystery Order Test Result for ${data.sellerName}. Please view this email in HTML format for complete information.`, // Fallback text
          {
            htmlBody: htmlForEmail,
            name: CONFIG.EMAIL_SENDER_NAME,     // The name displayed as the sender
            replyTo: CONFIG.EMAIL_SENDER_ADDRESS // The reply-to address
          }
        );
        Logger.log(`Email sent successfully to ${email} (${data.sellerName})`);
        successCount++;
      } else {
        Logger.log(`Invalid email address skipped: ${email}`);
      }
    }
    
    return successCount > 0; // Return true if at least one email was sent successfully
    
  } catch (error) {
    Logger.log(`Error sending email for ${data.sellerName}: ${error.message}`);
    return false;
  }
}

/**
 * Generate a tracking pixel for email open tracking
 * @param {string} sellerEmail The email address of the recipient
 * @param {string} sellerId The ID of the seller
 * @return {string} HTML for a tracking pixel
 */
function generateTrackingPixel(sellerEmail, sellerId) {
  // Create a unique tracking ID
  const trackingId = Utilities.getUuid();
  
  // Store this tracking ID in a separate sheet for later reference
  recordTrackingId(trackingId, sellerEmail, sellerId);
  
  // Build the tracking URL
  const trackingUrl = `${CONFIG.WEB_APP_URL}?id=${trackingId}&action=open`;
  
  // Return HTML for a 1x1 transparent pixel with the tracking URL
  return `<img src="${trackingUrl}" width="1" height="1" alt="" style="display:none">`;
}

/**
 * Record the tracking ID and email details in a separate sheet
 * @param {string} trackingId The unique tracking ID
 * @param {string} email The recipient email
 * @param {string} sellerId The seller ID
 */
function recordTrackingId(trackingId, email, sellerId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try to get the tracking sheet, or create it if it doesn't exist
    let trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    if (!trackingSheet) {
      trackingSheet = spreadsheet.insertSheet(CONFIG.TRACKING_SHEET_NAME);
      // Add headers
      trackingSheet.appendRow(['Tracking ID', 'Email', 'Seller ID', 'Send Date', 'Open Date', 'Opened', 'Views']);
    }
    
    // Add the new tracking record
    trackingSheet.appendRow([
      trackingId,
      email,
      sellerId,
      new Date(),
      '',
      'No',
      0 // Initialize Views count to 0
    ]);
    
    Logger.log(`Tracking ID ${trackingId} recorded for ${email}`);
  } catch (error) {
    Logger.log(`Error recording tracking ID: ${error.message}`);
  }
}

/**
 * Get email statistics from the tracking sheet
 * @return {Object} Statistics about email opens and engagement
 */
function getEmailStats() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    
    if (!trackingSheet) {
      return { error: 'Tracking sheet not found. It will be created automatically after emails are sent.' };
    }
    
    const data = trackingSheet.getDataRange().getValues();
    
    // Skip header row
    if (data.length <= 1) {
      return { error: 'No tracking data available. Statistics will appear after emails are sent.' };
    }
    
    let totalEmails = data.length - 1;
    let openedEmails = 0;
    let totalViews = 0;
    let passedResults = 0;
    let failedResults = 0;
    
    // Track date statistics
    const today = new Date();
    let lastWeekEmails = 0;
    let lastWeekOpened = 0;
    const oneWeekAgo = new Date(today);
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
    
    for (let i = 1; i < data.length; i++) {
      const sendDate = new Date(data[i][3]);
      
      // Check if email was opened
      if (data[i][5] === 'Yes') {
        openedEmails++;
        totalViews += Number(data[i][6] || 1);
      }
      
      // Count emails from last week
      if (sendDate > oneWeekAgo) {
        lastWeekEmails++;
        if (data[i][5] === 'Yes') {
          lastWeekOpened++;
        }
      }
    }
    
    const openRate = totalEmails > 0 ? (openedEmails / totalEmails * 100).toFixed(2) : 0;
    const lastWeekOpenRate = lastWeekEmails > 0 ? (lastWeekOpened / lastWeekEmails * 100).toFixed(2) : 0;
    const avgViews = openedEmails > 0 ? (totalViews / openedEmails).toFixed(2) : 0;
    
    return {
      totalEmails: totalEmails,
      openedEmails: openedEmails,
      openRate: `${openRate}%`,
      totalViews: totalViews,
      averageViews: avgViews,
      lastWeekEmails: lastWeekEmails,
      lastWeekOpened: lastWeekOpened,
      lastWeekOpenRate: `${lastWeekOpenRate}%`
    };
  } catch (error) {
    Logger.log(`Error getting email stats: ${error.message}`);
    return { error: error.message };
  }
}

/**
 * Set up a weekly trigger to run every Monday
 */
function setWeeklySchedule() {
  try {
    // Delete any existing triggers with the same function name
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'main') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // Create a new trigger to run every Monday at 9:00 AM
    ScriptApp.newTrigger('main')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
    
    Logger.log('Weekly trigger created to run at 9:00 AM every Monday');
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Success', 'Weekly schedule set! The script will run automatically at 9:00 AM every Monday.', ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the message
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error setting up trigger: ${error.message}`);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error', `Could not set up weekly schedule: ${error.message}`, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the error
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    return false;
  }
}
