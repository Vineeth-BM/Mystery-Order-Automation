/**
 * Mystery Order Email Automation - Main Controller
 * This script connects all components and provides the main entry points
 * @version 1.0
 * @lastModified 2025-04-18
 */

/**
 * Main function to send mystery order result notifications
 * Can be triggered manually or via time-based trigger
 */
function main() {
  try {
    // Start logging
    Logger.log('Starting mystery order notification process');
    
    // Access the sheet and get data
    const data = getSheetData();
    if (!data || data.length <= CONFIG.HEADER_ROW_COUNT) {
      Logger.log('No data found or only header row exists');
      
      // Check if we're running from UI
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Warning', 'No data found or only header row exists.', ui.ButtonSet.OK);
      } catch (e) {
        // Running from trigger, just log the error
        Logger.log('Not displaying UI alert: running from trigger');
      }
      
      return;
    }
    
    // Process the data and send emails
    const stats = processData(data);
    
    // Log summary statistics
    Logger.log(`Process completed. Results: ${stats.processed} rows processed, ${stats.emailsSent} emails sent, ${stats.errors} errors, ${stats.skippedNoEmail} skipped due to missing email`);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Success', 
              `Process completed:\n${stats.processed} rows processed\n${stats.emailsSent} emails sent\n${stats.errors} errors\n${stats.skippedNoEmail} skipped (no email)`, 
              ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the message
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    return stats;
    
  } catch (error) {
    Logger.log(`Critical error in main function: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Check if we're running from UI before showing alert
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error', `An error occurred: ${error.message}`, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the error
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    throw error; // Rethrow to see in Apps Script logs
  }
}

/**
 * Processes data rows and sends emails for mystery order results
 * @param {Array} data 2D array of sheet data
 * @return {Object} Statistics about the process
 */
function processData(data) {
  // Statistics tracking
  const stats = {
    processed: 0,
    emailsSent: 0,
    errors: 0,
    skippedNoEmail: 0
  };
  
  // Create template objects for dynamically constructing HTML
  let passedTemplate, failedTemplate;
  try {
    // Load Japanese templates
    passedTemplate = HtmlService.createTemplateFromFile(CONFIG.PASSED_EMAIL_TEMPLATE);
    failedTemplate = HtmlService.createTemplateFromFile(CONFIG.FAILED_EMAIL_TEMPLATE);
    
    Logger.log('Email templates loaded successfully');
  } catch (error) {
    Logger.log(`Error loading email templates: ${error.message}`);
    stats.errors++;
    return stats;
  }
  
  // Iterate over each row of data (skip the header row)
  for (let i = CONFIG.HEADER_ROW_COUNT; i < data.length; i++) {
    try {
      stats.processed++;
      
      // Skip rows that don't have enough columns
      if (data[i].length <= COLUMNS.REPORT) {
        Logger.log(`Row ${i+1}: Not enough columns (${data[i].length} found)`);
        continue;
      }
      
      // Extract row data
      const rowData = extractRowData(data[i]);
      
      // Skip if no email is available
      if (!rowData.sellerEmail) {
        Logger.log(`Row ${i+1}: No email found for seller ${rowData.sellerName} (ID: ${rowData.sellerId})`);
        stats.skippedNoEmail++;
        continue;
      }

      // Skip if no result is available
      if (!rowData.result) {
        Logger.log(`Row ${i+1}: No result found for seller ${rowData.sellerName} (ID: ${rowData.sellerId})`);
        continue;
      }
      
      // Send the appropriate email based on result
      let emailSent = false;
      
      if (rowData.result.toLowerCase() === 'passed') {
        emailSent = sendNotificationEmail(passedTemplate, rowData, CONFIG.EMAIL_SUBJECT_PASSED);
        Logger.log(`Row ${i+1}: Sending passed template to ${rowData.sellerName}`);
      } else if (rowData.result.toLowerCase() === 'failed') {
        emailSent = sendNotificationEmail(failedTemplate, rowData, CONFIG.EMAIL_SUBJECT_FAILED);
        Logger.log(`Row ${i+1}: Sending failed template to ${rowData.sellerName}`);
      } else {
        Logger.log(`Row ${i+1}: Unknown result "${rowData.result}" for seller ${rowData.sellerName}`);
        continue;
      }
      
      if (emailSent) {
        stats.emailsSent++;
      } else {
        stats.errors++;
      }
      
    } catch (error) {
      Logger.log(`Error processing row ${i+1}: ${error.message}`);
      stats.errors++;
      continue; // Continue with the next row
    }
  }
  
  return stats;
}

/**
 * Test the system
 */
function testSystem() {
  try {
    let results = [];
    
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    results.push(`✓ Spreadsheet access OK: "${spreadsheet.getName()}"`);
    
    // Test mystery order sheet access
    const sheet = spreadsheet.getSheetByName(CONFIG.MYSTERY_SHEET_NAME);
    if (sheet) {
      results.push(`✓ Mystery order sheet access OK: "${CONFIG.MYSTERY_SHEET_NAME}"`);
      
      // Test data access
      const data = sheet.getDataRange().getValues();
      results.push(`✓ Data access OK: ${data.length} rows retrieved`);
      
      // Count rows with valid data
      if (data.length > CONFIG.HEADER_ROW_COUNT) {
        let passedCount = 0;
        let failedCount = 0;
        let emailCount = 0;
        
        for (let i = CONFIG.HEADER_ROW_COUNT; i < data.length; i++) {
          if (data[i].length > COLUMNS.REPORT) {
            const result = data[i][COLUMNS.RESULT];
            const hasEmail = Boolean(data[i][COLUMNS.EMAIL]);
            
            if (result) {
              if (result.toLowerCase() === 'passed') {
                passedCount++;
              } else if (result.toLowerCase() === 'failed') {
                failedCount++;
              }
              
              if (hasEmail) {
                emailCount++;
              }
            }
          }
        }
        results.push(`${passedCount} rows with 'passed' result, ${failedCount} with 'failed' result`);
        results.push(`${emailCount} rows with valid emails out of ${data.length - CONFIG.HEADER_ROW_COUNT} total entries`);
      }
    } else {
      results.push(`✗ Mystery order sheet not found: "${CONFIG.MYSTERY_SHEET_NAME}"`);
    }
    
    // Test email template access
    try {
      const passedTemplate = HtmlService.createTemplateFromFile(CONFIG.PASSED_EMAIL_TEMPLATE);
      results.push(`✓ Passed email template access OK: "${CONFIG.PASSED_EMAIL_TEMPLATE}"`);
    } catch (e) {
      results.push(`✗ Passed email template access failed: "${CONFIG.PASSED_EMAIL_TEMPLATE}"`);
    }
    
    try {
      const failedTemplate = HtmlService.createTemplateFromFile(CONFIG.FAILED_EMAIL_TEMPLATE);
      results.push(`✓ Failed email template access OK: "${CONFIG.FAILED_EMAIL_TEMPLATE}"`);
    } catch (e) {
      results.push(`✗ Failed email template access failed: "${CONFIG.FAILED_EMAIL_TEMPLATE}"`);
    }
    
    // Check Gmail access
    try {
      // Just checking if we can access Gmail without actually sending
      const quota = MailApp.getRemainingDailyQuota();
      results.push(`✓ Email sending available (${quota} emails remaining in daily quota)`);
    } catch (e) {
      results.push(`✗ Email sending access failed: ${e.message}`);
    }
    
    // Log results
    const resultsText = results.join('\n');
    Logger.log('Test completed:\n' + resultsText);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('System Test Results', resultsText, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the message
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    return true;
  } catch (error) {
    Logger.log(`Test failed: ${error.message}`);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Test Failed', `Error: ${error.message}`, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the error
      Logger.log('Not displaying UI alert: running from trigger');
    }
    
    return false;
  }
}
