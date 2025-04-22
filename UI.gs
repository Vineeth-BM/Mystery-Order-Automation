/**
 * Mystery Order Email Automation - UI Components
 * User interface functions and controls
 * @version 1.0
 * @lastModified 2025-04-18
 */

/**
 * Creates a custom menu in the Google Sheets UI when the spreadsheet is opened
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

     // Create a menu
    ui.createMenu('Mystery Order Emails')
      .addItem('Send Mystery Order Emails Now', 'main')
      .addSeparator()
      .addItem('Test System', 'testSystem')
      .addItem('Set Weekly Schedule', 'setWeeklySchedule')
      .addSeparator()
      .addItem('Show Email Statistics', 'showEmailStats')
      .addToUi();
    
    Logger.log('Custom menu created successfully');
  } catch (error) {
    Logger.log(`Error creating custom menu: ${error.message}`);
  }
}

/**
 * Show email statistics to the user
 */
function showEmailStats() {
  try {
    const stats = getEmailStats();
    
    let message;
    if (stats.error) {
      message = `Error: ${stats.error}`;
    } else {
      message = `Email Statistics:\n\n` +
                `Total Emails Sent: ${stats.totalEmails}\n` +
                `Emails Opened: ${stats.openedEmails}\n` +
                `Open Rate: ${stats.openRate}\n` +
                `Total Views: ${stats.totalViews}\n` +
                `Average Views Per Opened Email: ${stats.averageViews}\n\n` +
                `Last 7 Days Emails: ${stats.lastWeekEmails}\n` +
                `Last 7 Days Opens: ${stats.lastWeekOpened}\n` +
                `Last 7 Days Open Rate: ${stats.lastWeekOpenRate}`;
    }
    
    Logger.log('Email statistics:\n' + message);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Email Statistics', message, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the message
      Logger.log('Not displaying UI alert: running from trigger');
    }
  } catch (error) {
    Logger.log(`Error showing email stats: ${error.message}`);
    
    // Check if we're running from UI
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error', `Failed to get email statistics: ${error.message}`, ui.ButtonSet.OK);
    } catch (e) {
      // Running from trigger, just log the error
      Logger.log('Not displaying UI alert: running from trigger');
    }
  }
}
