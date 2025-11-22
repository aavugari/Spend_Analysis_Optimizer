// ============================================================================
// TRANSACTION MERGER - REFACTORED
// ============================================================================

/**
 * Configuration for merger (move these to PropertiesService for security)
 */
const MERGER_CONFIG = {
  // TODO: Move these to PropertiesService.getScriptProperties()
  SURYA_SHEET: {
    url: "YOUR_SURYA_SHEET_URL_HERE",
    sheetName: "Sheet1",
    sourceLabel: "Surya"
  },
  NAMITA_SHEET: {
    url: "YOUR_NAMITA_SHEET_URL_HERE", 
    sheetName: "Sheet1",
    sourceLabel: "Wife"
  },
  MASTER_SHEET_NAME: "Master",
  HEADERS_WITH_SOURCE: ["Bank", "Date", "Amount", "Transaction Info", "Transaction Type", "Category", "Card Last 4", "Month", "Year", "Source"]
};

/**
 * Main function to merge transaction sheets
 */
function mergeTransactionSheets() {
  try {
    Logger.log("🔄 Starting transaction merge process");
    
    var masterSheet = initializeMasterSheet();
    var mergedCounts = {
      "Surya": 0,
      "Wife": 0,
      "Total": 0
    };
    
    // Merge Surya's data
    mergedCounts["Surya"] = mergeDataFromSheet(
      MERGER_CONFIG.SURYA_SHEET.url,
      MERGER_CONFIG.SURYA_SHEET.sheetName,
      MERGER_CONFIG.SURYA_SHEET.sourceLabel,
      masterSheet
    );
    
    // Merge Namita's data
    mergedCounts["Wife"] = mergeDataFromSheet(
      MERGER_CONFIG.NAMITA_SHEET.url,
      MERGER_CONFIG.NAMITA_SHEET.sheetName,
      MERGER_CONFIG.NAMITA_SHEET.sourceLabel,
      masterSheet
    );
    
    mergedCounts["Total"] = mergedCounts["Surya"] + mergedCounts["Wife"];
    
    Logger.log("✅ TRANSACTION MERGER completed: Surya=" + mergedCounts["Surya"] + ", Wife=" + mergedCounts["Wife"] + ", Total=" + mergedCounts["Total"]);
    
  } catch (error) {
    Logger.log("❌ CRITICAL ERROR in mergeTransactionSheets: " + error.toString());
    throw error;
  }
}

/**
 * Initialize or get master sheet
 */
function initializeMasterSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = spreadsheet.getSheetByName(MERGER_CONFIG.MASTER_SHEET_NAME);
    
    if (!masterSheet) {
      masterSheet = spreadsheet.insertSheet(MERGER_CONFIG.MASTER_SHEET_NAME);
      Logger.log("✅ Created new Master sheet");
    }
    
    // Clear existing data
    masterSheet.clear();
    
    // Add headers
    masterSheet.appendRow(MERGER_CONFIG.HEADERS_WITH_SOURCE);
    Logger.log("✅ Master sheet initialized with headers");
    
    return masterSheet;
    
  } catch (error) {
    Logger.log("❌ Error initializing master sheet: " + error.toString());
    throw error;
  }
}

/**
 * Merge data from a specific sheet
 */
function mergeDataFromSheet(sheetUrl, sheetName, sourceLabel, masterSheet) {
  try {
    Logger.log("📊 Merging data from: " + sourceLabel);
    
    // Open source spreadsheet
    var sourceSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    
    if (!sourceSheet) {
      Logger.log("⚠️ Sheet '" + sheetName + "' not found in " + sourceLabel + "'s spreadsheet");
      return 0;
    }
    
    var lastRow = sourceSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("ℹ️ No data found in " + sourceLabel + "'s sheet");
      return 0;
    }
    
    // Get data (skip header row)
    var sourceData = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
    
    // Add source label to each row and validate
    var processedData = [];
    var validCount = 0;
    
    sourceData.forEach(function(row, index) {
      try {
        // Basic validation
        if (isValidTransactionRow(row)) {
          // Add source label
          var rowWithSource = row.slice(); // Create copy
          rowWithSource.push(sourceLabel);
          processedData.push(rowWithSource);
          validCount++;
        } else {
          Logger.log("⚠️ Skipping invalid row " + (index + 2) + " from " + sourceLabel);
        }
      } catch (error) {
        Logger.log("⚠️ Error processing row " + (index + 2) + " from " + sourceLabel + ": " + error.toString());
      }
    });
    
    // Append to master sheet
    if (processedData.length > 0) {
      var targetRange = masterSheet.getRange(
        masterSheet.getLastRow() + 1, 
        1, 
        processedData.length, 
        processedData[0].length
      );
      targetRange.setValues(processedData);
      
      Logger.log("✅ Added " + validCount + " transactions from " + sourceLabel);
    }
    
    return validCount;
    
  } catch (error) {
    Logger.log("❌ Error merging data from " + sourceLabel + ": " + error.toString());
    return 0;
  }
}

/**
 * Validate transaction row
 */
function isValidTransactionRow(row) {
  // Check if row has minimum required fields (now 9 columns)
  if (!row || row.length < 9) return false;
  
  // Check if essential fields are not empty
  var bank = row[0];
  var date = row[1];
  var amount = row[2];
  
  if (!bank || !date || !amount) return false;
  if (isNaN(amount) || amount <= 0) return false;
  
  return true;
}

// ============================================================================
// TELEGRAM NOTIFICATION SYSTEM - REFACTORED
// ============================================================================

/**
 * Send daily summary via Telegram with improved error handling
 */
function sendDailySummaryTelegram() {
  try {
    Logger.log("📱 Preparing Telegram daily summary");
    
    var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MERGER_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) {
      Logger.log("❌ Master sheet not found. Run mergeTransactionSheets() first.");
      return;
    }
    
    var summaryData = calculateDailySummary(masterSheet);
    
    if (summaryData.totalToday === 0 && summaryData.totalMTD === 0) {
      Logger.log("ℹ️ No transactions to report");
      return;
    }
    
    var message = buildTelegramMessage(summaryData);
    sendTelegramMessage(message);
    
  } catch (error) {
    Logger.log("❌ Error sending Telegram summary: " + error.toString());
  }
}

/**
 * Calculate daily and MTD summaries
 */
function calculateDailySummary(masterSheet) {
  var today = new Date();
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
  var monthStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM yyyy");
  
  var lastRow = masterSheet.getLastRow();
  if (lastRow <= 1) {
    return { totalToday: 0, totalMTD: 0, todayBreakdown: {}, mtdBreakdown: {} };
  }
  
  var data = masterSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  var todayBreakdown = {};
  var mtdBreakdown = {};
  var totalToday = 0;
  var totalMTD = 0;
  
  data.forEach(function(row) {
    try {
      var bank = row[0];
      var date = new Date(row[1]);
      var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
      var monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy");
      var amount = parseFloat(row[2]);
      var type = row[4];
      var person = row[9]; // Source column (now at index 9)
      
      // Only process debit transactions
      if (type === "Debit" && !isNaN(amount)) {
        // Today's totals
        if (dateStr === todayStr) {
          if (!todayBreakdown[person]) todayBreakdown[person] = {};
          if (!todayBreakdown[person][bank]) todayBreakdown[person][bank] = 0;
          todayBreakdown[person][bank] += amount;
          totalToday += amount;
        }
        
        // Month-to-date totals
        if (monthYear === monthStr) {
          if (!mtdBreakdown[person]) mtdBreakdown[person] = {};
          if (!mtdBreakdown[person][bank]) mtdBreakdown[person][bank] = 0;
          mtdBreakdown[person][bank] += amount;
          totalMTD += amount;
        }
      }
    } catch (error) {
      Logger.log("⚠️ Error processing row for summary: " + error.toString());
    }
  });
  
  return {
    todayStr: todayStr,
    monthStr: monthStr,
    totalToday: totalToday,
    totalMTD: totalMTD,
    todayBreakdown: todayBreakdown,
    mtdBreakdown: mtdBreakdown
  };
}

/**
 * Build formatted Telegram message
 */
function buildTelegramMessage(summaryData) {
  var message = "📅 *Daily Spend Summary* - " + summaryData.todayStr + "\\n\\n";
  
  // Today's spending
  if (summaryData.totalToday > 0) {
    for (var person in summaryData.todayBreakdown) {
      message += "👤 *" + person + "*\\n";
      for (var bank in summaryData.todayBreakdown[person]) {
        message += "💳 " + bank + ": Rs. " + summaryData.todayBreakdown[person][bank].toFixed(2) + "\\n";
      }
      message += "\\n";
    }
    message += "📊 *Total Spent Today*: Rs. " + summaryData.totalToday.toFixed(2) + "\\n\\n";
  } else {
    message += "No spends today ✅\\n\\n";
  }
  
  // Month-to-date spending
  message += "📆 *Month-to-Date (" + summaryData.monthStr + ")*\\n\\n";
  for (var person in summaryData.mtdBreakdown) {
    message += "👤 *" + person + "*\\n";
    for (var bank in summaryData.mtdBreakdown[person]) {
      message += "💳 " + bank + ": Rs. " + summaryData.mtdBreakdown[person][bank].toFixed(2) + "\\n";
    }
    message += "\\n";
  }
  message += "📊 *Total MTD*: Rs. " + summaryData.totalMTD.toFixed(2);
  
  return message;
}

/**
 * Send message to Telegram with secure token handling
 */
function sendTelegramMessage(message) {
  try {
    // TODO: Move these to PropertiesService for security
    var token = "YOUR_TELEGRAM_BOT_TOKEN_HERE"; // SECURITY: Move to PropertiesService
    var chatId = "YOUR_TELEGRAM_CHAT_ID_HERE"; // SECURITY: Move to PropertiesService
    
    var url = "https://api.telegram.org/bot" + token + "/sendMessage";
    var payload = {
      chat_id: chatId,
      text: message,
      parse_mode: "Markdown"
    };
    
    var options = {
      method: "post",
      payload: payload
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseData = JSON.parse(response.getContentText());
    
    if (responseData.ok) {
      Logger.log("✅ Telegram message sent successfully");
    } else {
      Logger.log("❌ Telegram API error: " + responseData.description);
    }
    
  } catch (error) {
    Logger.log("❌ Error sending Telegram message: " + error.toString());
  }
}

// ============================================================================
// SECURITY IMPROVEMENT FUNCTIONS
// ============================================================================

/**
 * Setup secure configuration (run once to store secrets)
 */
function setupSecureConfig() {
  var properties = PropertiesService.getScriptProperties();
  
  // Store Telegram credentials securely
  properties.setProperties({
    'TELEGRAM_BOT_TOKEN': 'YOUR_TELEGRAM_BOT_TOKEN_HERE',
    'TELEGRAM_CHAT_ID': 'YOUR_TELEGRAM_CHAT_ID_HERE',
    'SURYA_SHEET_URL': 'YOUR_SURYA_SHEET_URL_HERE',
    'NAMITA_SHEET_URL': 'YOUR_NAMITA_SHEET_URL_HERE'
  });
  
  Logger.log("✅ Secure configuration stored");
}

/**
 * Get secure configuration
 */
function getSecureConfig() {
  var properties = PropertiesService.getScriptProperties();
  return {
    telegramToken: properties.getProperty('TELEGRAM_BOT_TOKEN'),
    telegramChatId: properties.getProperty('TELEGRAM_CHAT_ID'),
    suryaSheetUrl: properties.getProperty('SURYA_SHEET_URL'),
    namitaSheetUrl: properties.getProperty('NAMITA_SHEET_URL')
  };
}