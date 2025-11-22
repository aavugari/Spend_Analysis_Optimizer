// ============================================================================
// CONFIGURATION TEMPLATE
// ============================================================================
// Copy this file to config.js and fill in your actual values
// Add config.js to .gitignore to keep your credentials secure

/**
 * Configuration template for Spend Analysis Optimizer
 * Replace placeholder values with your actual credentials and URLs
 */

const CONFIG_TEMPLATE = {
  // Google Sheets URLs
  SURYA_SHEET_URL: "https://docs.google.com/spreadsheets/d/YOUR_SURYA_SHEET_ID/edit",
  NAMITA_SHEET_URL: "https://docs.google.com/spreadsheets/d/YOUR_NAMITA_SHEET_ID/edit",
  
  // Telegram Bot Configuration
  TELEGRAM_BOT_TOKEN: "YOUR_BOT_TOKEN:FROM_BOTFATHER",
  TELEGRAM_CHAT_ID: "YOUR_CHAT_ID_NUMBER",
  
  // Sheet Names (usually don't need to change)
  SHEET_NAME: "Sheet1",
  MASTER_SHEET_NAME: "Master"
};

/**
 * How to get your configuration values:
 * 
 * 1. Google Sheets URLs:
 *    - Create your Google Sheets for transaction data
 *    - Copy the full URL from your browser
 *    - Replace YOUR_SURYA_SHEET_ID and YOUR_NAMITA_SHEET_ID
 * 
 * 2. Telegram Bot Token:
 *    - Message @BotFather on Telegram
 *    - Create a new bot with /newbot command
 *    - Copy the token provided
 * 
 * 3. Telegram Chat ID:
 *    - Message your bot first
 *    - Visit: https://api.telegram.org/botYOUR_BOT_TOKEN/getUpdates
 *    - Find your chat ID in the response
 * 
 * 4. Setup in Google Apps Script:
 *    - Run setupSecureConfig() once to store credentials securely
 *    - Update MERGER_CONFIG URLs in Transaction-Merger.js
 */

// Example usage in your scripts:
// Replace hardcoded values with: getSecureConfig().telegramToken