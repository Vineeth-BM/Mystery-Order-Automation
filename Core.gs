/**
 * Mystery Order Email Automation - Core Components
 * Core functionality and configuration for the mystery order email system
 * @version 1.0
 * @lastModified 2025-04-18
 */

// Configuration constants - edit these as needed
const CONFIG = {
  MYSTERY_SHEET_NAME: 'Mystery Order',
  // Email templates (Japanese only)
  PASSED_EMAIL_TEMPLATE: 'passed_email_template_jp',
  FAILED_EMAIL_TEMPLATE: 'failed_email_template_jp',
  // Email subjects (Japanese)
  EMAIL_SUBJECT_PASSED: 'ミステリーオーダーテスト：合格しました',
  EMAIL_SUBJECT_FAILED: 'ミステリーオーダーテスト：改善が必要です',
  // Sheet settings
  HEADER_ROW_COUNT: 1,
  DEBUG_MODE: true, // Set to true for verbose logging
  // Email configuration
  EMAIL_SENDER_NAME: 'Back Market Quality Team',
  EMAIL_SENDER_ADDRESS: 'yuho.suzuki@backmarket.com', // This will be used as the reply-to address
  // Tracking configuration
  TRACKING_SHEET_NAME: 'Mystery Email Tracking',
  // Web app URL for tracking pixel
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbyZTRbYV2UJz0MtQWIOS4uilC87fUWuOUoNM5Lj7QOiR3dP5cUqviREKN9FbH6xeBg-PQ/exec',
  // Resources URLs
  SELLER_SUPPORT_CENTER_URL: 'https://drive.google.com/drive/u/0/folders/1UvnSv6Knxu0-VHu01KFn7MP8A8pYmML4',
  QUALITY_DASHBOARD_URL: 'https://www.backmarket.fr/bo_merchant/insights?tableau=quality-dashboard'
};

// Column indices (0 based) - matching the Mystery Order sheet
const COLUMNS = {
  SELLER_ID: 0,
  SELLER_NAME: 1,
  EMAIL: 2,
  RESULT: 3,
  REPORT: 4
};

/**
 * Retrieves data from the specified sheet
 * @return {Array} 2D array of sheet data
 */
function getSheetData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Accessing spreadsheet: "${spreadsheet.getName()}"`);
    
    const sheet = spreadsheet.getSheetByName(CONFIG.MYSTERY_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${CONFIG.MYSTERY_SHEET_NAME}" not found!`);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log(`Retrieved ${data.length} rows from sheet "${CONFIG.MYSTERY_SHEET_NAME}"`);
    
    return data;
  } catch (error) {
    Logger.log(`Error retrieving sheet data: ${error.message}`);
    return null;
  }
}

/**
 * Extracts and formats data from a single row
 * @param {Array} row Single row of data from the sheet
 * @return {Object} Formatted data object
 */
function extractRowData(row) {
  return {
    sellerId: row[COLUMNS.SELLER_ID] || '',
    sellerName: row[COLUMNS.SELLER_NAME] || '',
    sellerEmail: row[COLUMNS.EMAIL] || '',
    result: row[COLUMNS.RESULT] || '',
    reportLink: row[COLUMNS.REPORT] || '',
    
    // Resource URLs
    sellerSupportCenterUrl: CONFIG.SELLER_SUPPORT_CENTER_URL,
    qualityDashboardUrl: CONFIG.QUALITY_DASHBOARD_URL
  };
}

/**
 * Validates an email address format
 * @param {String} email The email address to validate
 * @return {Boolean} Whether the email is valid
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  
  // Basic email validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
