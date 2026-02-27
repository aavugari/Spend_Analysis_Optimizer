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
    sourceLabel: "Surya (Husband)"
  },
  NAMITA_SHEET: {
    url: "YOUR_NAMITA_SHEET_URL_HERE", 
    sheetName: "Sheet1",
    sourceLabel: "Namita (Wife)"
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
      "Surya (Husband)": 0,
      "Namita (Wife)": 0,
      "Total": 0
    };
    
    // Merge Surya (Husband)'s data
    mergedCounts["Surya (Husband)"] = mergeDataFromSheet(
      MERGER_CONFIG.SURYA_SHEET.url,
      MERGER_CONFIG.SURYA_SHEET.sheetName,
      MERGER_CONFIG.SURYA_SHEET.sourceLabel,
      masterSheet
    );
    
    // Merge Namita (Wife)'s data
    mergedCounts["Namita (Wife)"] = mergeDataFromSheet(
      MERGER_CONFIG.NAMITA_SHEET.url,
      MERGER_CONFIG.NAMITA_SHEET.sheetName,
      MERGER_CONFIG.NAMITA_SHEET.sourceLabel,
      masterSheet
    );
    
    mergedCounts["Total"] = mergedCounts["Surya (Husband)"] + mergedCounts["Namita (Wife)"];
    
    Logger.log(
      "✅ TRANSACTION MERGER completed: Surya (Husband)=" +
      mergedCounts["Surya (Husband)"] +
      ", Namita (Wife)=" +
      mergedCounts["Namita (Wife)"] +
      ", Total=" +
      mergedCounts["Total"]
    );
    
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
// SMART GOAL TRACKING & ALERT SYSTEM
// ============================================================================

/**
 * Categories to exclude from goal tracking (subscriptions and others)
 */
function getExcludedCategories() {
  return [
    "Others",
    "Google Subscription", 
    "Youtube Subscription",
    "Netflix",
    "Subscription",
    "Apple"
  ];
}

/**
 * Check if category should be excluded from goal tracking
 */
function isSubscriptionCategory(category) {
  var excludedCategories = getExcludedCategories();
  return excludedCategories.some(function(excluded) {
    return category && category.toLowerCase().includes(excluded.toLowerCase());
  });
}

/**
 * Calculate 6-month average and set progressive reduction goals
 */
function calculateGoalTargets(masterSheet) {
  var sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
  
  var data = masterSheet.getDataRange().getValues();
  var categoryTotals = {};
  var monthCount = {};
  
  // Calculate 6-month averages by category (excluding subscriptions)
  data.slice(1).forEach(function(row) {
    var date = new Date(row[1]);
    var amount = parseFloat(row[2]);
    var category = row[5];
    var type = row[4];
    
    if (date >= sixMonthsAgo && type === "Debit" && !isNaN(amount) && !isSubscriptionCategory(category)) {
      var monthKey = date.getFullYear() + "-" + date.getMonth();
      
      if (!categoryTotals[category]) categoryTotals[category] = {};
      if (!categoryTotals[category][monthKey]) categoryTotals[category][monthKey] = 0;
      categoryTotals[category][monthKey] += amount;
      
      if (!monthCount[monthKey]) monthCount[monthKey] = true;
    }
  });
  
  var totalMonths = Object.keys(monthCount).length || 6;
  var goals = {};
  
  // Calculate progressive reduction goals
  for (var category in categoryTotals) {
    var monthlyTotals = Object.values(categoryTotals[category]);
    var average = monthlyTotals.reduce((sum, val) => sum + val, 0) / totalMonths;
    
    var currentMonth = getCurrentGoalPhase();
    var reductionRate = currentMonth <= 3 ? 0.10 : currentMonth <= 6 ? 0.20 : 0.30;
    
    goals[category] = {
      baseline: average,
      target: average * (1 - reductionRate),
      reductionRate: reductionRate * 100,
      phase: currentMonth <= 3 ? "Phase 1" : currentMonth <= 6 ? "Phase 2" : "Phase 3"
    };
  }
  
  return goals;
}

/**
 * Get current goal phase (1-9 months from start)
 */
function getCurrentGoalPhase() {
  var startDate = new Date("2025-01-01"); // Goal start date
  var currentDate = new Date();
  var monthsDiff = (currentDate.getFullYear() - startDate.getFullYear()) * 12 + 
                   (currentDate.getMonth() - startDate.getMonth()) + 1;
  return Math.min(monthsDiff, 9);
}

/**
 * Calculate current month spending by category (excluding subscriptions)
 */
function getCurrentMonthSpending(masterSheet) {
  var currentMonth = new Date().getMonth();
  var currentYear = new Date().getFullYear();
  var data = masterSheet.getDataRange().getValues();
  var spending = {};
  
  data.slice(1).forEach(function(row) {
    var date = new Date(row[1]);
    var amount = parseFloat(row[2]);
    var category = row[5];
    var type = row[4];
    
    if (date.getMonth() === currentMonth && date.getFullYear() === currentYear && 
        type === "Debit" && !isNaN(amount) && !isSubscriptionCategory(category)) {
      spending[category] = (spending[category] || 0) + amount;
    }
  });
  
  return spending;
}

/**
 * Calculate current month subscription spending
 */
function getCurrentMonthSubscriptions(masterSheet) {
  var currentMonth = new Date().getMonth();
  var currentYear = new Date().getFullYear();
  var data = masterSheet.getDataRange().getValues();
  var subscriptions = {};
  
  data.slice(1).forEach(function(row) {
    var date = new Date(row[1]);
    var amount = parseFloat(row[2]);
    var category = row[5];
    var info = row[3];
    var type = row[4];
    
    if (date.getMonth() === currentMonth && date.getFullYear() === currentYear && 
        type === "Debit" && !isNaN(amount) && isSubscriptionCategory(category)) {
      
      var serviceName = getServiceName(info, category);
      subscriptions[serviceName] = (subscriptions[serviceName] || 0) + amount;
    }
  });
  
  return subscriptions;
}

/**
 * Extract service name from transaction info
 */
function getServiceName(info, category) {
  if (info.toLowerCase().includes('netflix')) return 'Netflix';
  if (info.toLowerCase().includes('youtube') || info.toLowerCase().includes('google')) return 'YouTube Premium';
  if (info.toLowerCase().includes('apple')) return 'Apple Services';
  if (info.toLowerCase().includes('amazon')) return 'Amazon Prime';
  if (info.toLowerCase().includes('spotify')) return 'Spotify';
  
  return category || 'Other Subscription';
}

/**
 * Generate smart alerts and contextual messages
 */
function generateSmartAlerts(goals, currentSpending) {
  var alerts = [];
  var today = new Date();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var daysPassed = today.getDate();
  var monthProgress = daysPassed / daysInMonth;
  
  for (var category in goals) {
    var goal = goals[category];
    var spent = currentSpending[category] || 0;
    var targetSpent = goal.target * monthProgress;
    var budgetUsed = spent / goal.target;
    
    // Budget warning alerts (75% threshold)
    if (budgetUsed >= 0.75 && budgetUsed < 1.0) {
      var remaining = goal.target - spent;
      alerts.push({
        type: "warning",
        category: category,
        message: "⚠️ *" + category + "*: You're ₹" + remaining.toFixed(0) + " away from your goal! (" + (budgetUsed * 100).toFixed(0) + "% used)"
      });
    }
    
    // Budget exceeded alerts
    else if (budgetUsed >= 1.0) {
      var excess = spent - goal.target;
      alerts.push({
        type: "critical",
        category: category,
        message: "🚨 *" + category + "*: Budget exceeded by ₹" + excess.toFixed(0) + "! Consider reducing spending."
      });
    }
    
    // Unusual spending detection (spending too fast)
    else if (spent > targetSpent * 1.5 && daysPassed < 15) {
      alerts.push({
        type: "anomaly",
        category: category,
        message: "📊 *" + category + "*: Spending faster than usual. Current: ₹" + spent.toFixed(0) + ", Expected: ₹" + targetSpent.toFixed(0)
      });
    }
    
    // Positive progress alerts
    else if (budgetUsed < 0.5 && monthProgress > 0.5) {
      var saved = (goal.baseline - spent);
      if (saved > 0) {
        alerts.push({
          type: "positive",
          category: category,
          message: "🎉 *" + category + "*: Great progress! You've saved ₹" + saved.toFixed(0) + " compared to your baseline."
        });
      }
    }
  }
  
  return alerts;
}

/**
 * Generate contextual motivational messages
 */
function getContextualMessage(alerts, totalSavings) {
  var messages = [];
  
  if (totalSavings > 1000) {
    messages.push("💪 *Motivational*: Excellent! You've saved ₹" + totalSavings.toFixed(0) + " this month!");
  }
  
  var criticalAlerts = alerts.filter(a => a.type === "critical").length;
  var warningAlerts = alerts.filter(a => a.type === "warning").length;
  
  if (criticalAlerts === 0 && warningAlerts === 0) {
    messages.push("✅ *Educational*: All categories are on track! Keep up the great work!");
  } else if (criticalAlerts > 0) {
    messages.push("🎯 *Actionable*: Focus on reducing spending in " + criticalAlerts + " category(ies) to get back on track.");
  }
  
  // Add seasonal/contextual advice
  var month = new Date().getMonth();
  if (month === 10 || month === 11) { // Nov-Dec (Festival season)
    messages.push("🪔 *Seasonal*: Festival season - plan your spending to stay within goals!");
  }
  
  return messages;
}

// ============================================================================
// TELEGRAM NOTIFICATION SYSTEM - REFACTORED
// ============================================================================

/**
 * Send daily summary via Telegram with goal tracking and subscription tracking
 */
function sendDailySummaryTelegram() {
  try {
    Logger.log("📱 Preparing Telegram daily summary with goal tracking");
    
    var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MERGER_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) {
      Logger.log("❌ Master sheet not found. Run mergeTransactionSheets() first.");
      return;
    }
    
    var summaryData = calculateDailySummary(masterSheet);
    var goals = calculateGoalTargets(masterSheet);
    var currentSpending = getCurrentMonthSpending(masterSheet);
    var subscriptions = getCurrentMonthSubscriptions(masterSheet);
    var alerts = generateSmartAlerts(goals, currentSpending);
    
    var message = buildTelegramMessage(summaryData, goals, currentSpending, alerts, subscriptions);
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
 * Build formatted Telegram message with goal tracking and subscription tracking
 */
function buildTelegramMessage(summaryData, goals, currentSpending, alerts, subscriptions) {
  var message = "📅 *Daily Spend Summary* - " + summaryData.todayStr + "\n\n";
  
  // Today's spending
  if (summaryData.totalToday > 0) {
    for (var person in summaryData.todayBreakdown) {
      message += "👤 *" + person + "*\n";
      for (var bank in summaryData.todayBreakdown[person]) {
        message += "💳 " + bank + ": Rs. " + summaryData.todayBreakdown[person][bank].toFixed(2) + "\n";
      }
      message += "\n";
    }
    message += "📊 *Total Spent Today*: Rs. " + summaryData.totalToday.toFixed(2) + "\n\n";
  } else {
    message += "No spends today ✅\n\n";
  }
  
  // Month-to-date spending
  message += "📆 *Month-to-Date (" + summaryData.monthStr + ")*\n\n";
  for (var person in summaryData.mtdBreakdown) {
    message += "👤 *" + person + "*\n";
    for (var bank in summaryData.mtdBreakdown[person]) {
      message += "💳 " + bank + ": Rs. " + summaryData.mtdBreakdown[person][bank].toFixed(2) + "\n";
    }
    message += "\n";
  }
  message += "📊 *Total MTD*: Rs. " + summaryData.totalMTD.toFixed(2) + "\n\n";
  
  // Goal Progress Section
  if (goals && Object.keys(goals).length > 0) {
    message += "🎯 *Goal Progress*\n\n";
    var totalSavings = 0;
    
    for (var category in goals) {
      var goal = goals[category];
      var spent = currentSpending[category] || 0;
      var progress = Math.min((spent / goal.target) * 100, 100);
      var savings = goal.baseline - spent;
      
      if (savings > 0) totalSavings += savings;
      
      var emoji = progress < 75 ? "🟢" : progress < 100 ? "🟡" : "🔴";
      message += emoji + " *" + category + "*: " + progress.toFixed(0) + "% (₹" + spent.toFixed(0) + "/₹" + goal.target.toFixed(0) + ")\n";
    }
    
    if (totalSavings > 0) {
      message += "\n💰 *Total Savings*: ₹" + totalSavings.toFixed(0) + "\n";
    }
    message += "\n";
  }
  
  // Active Subscriptions Section
  if (subscriptions && Object.keys(subscriptions).length > 0) {
    message += "📱 *Active Subscriptions*\n";
    var totalSubscriptions = 0;
    
    for (var service in subscriptions) {
      var amount = subscriptions[service];
      totalSubscriptions += amount;
      message += "🔔 " + service + ": ₹" + amount.toFixed(2) + "\n";
    }
    
    message += "💳 *Total Subscriptions*: ₹" + totalSubscriptions.toFixed(2) + "\n\n";
  }
  
  // Smart Alerts Section
  if (alerts && alerts.length > 0) {
    message += "🚨 *Smart Alerts*\n\n";
    alerts.slice(0, 3).forEach(function(alert) {
      message += alert.message + "\n\n";
    });
  }
  
  // Contextual Messages
  var contextualMessages = getContextualMessage(alerts || [], totalSavings || 0);
  if (contextualMessages.length > 0) {
    contextualMessages.forEach(function(msg) {
      message += msg + "\n\n";
    });
  }
  
  return message;
}

/**
 * Send weekly summary via Telegram with goal tracking
 */
function sendWeeklySummaryTelegram() {
  try {
    Logger.log("📱 Preparing Telegram weekly summary with goal tracking");
    
    var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MERGER_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) {
      Logger.log("❌ Master sheet not found. Run mergeTransactionSheets() first.");
      return;
    }
    
    var weeklyData = calculateWeeklySummary(masterSheet);
    var goals = calculateGoalTargets(masterSheet);
    var currentSpending = getCurrentMonthSpending(masterSheet);
    var subscriptions = getCurrentMonthSubscriptions(masterSheet);
    var alerts = generateSmartAlerts(goals, currentSpending);
    
    var message = buildWeeklyTelegramMessage(weeklyData, goals, currentSpending, alerts, subscriptions);
    sendTelegramMessage(message);
    
  } catch (error) {
    Logger.log("❌ Error sending weekly Telegram summary: " + error.toString());
  }
}

/**
 * Calculate weekly summary data
 */
function calculateWeeklySummary(masterSheet) {
  var today = new Date();
  var weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  var data = masterSheet.getDataRange().getValues();
  var weeklySpending = {};
  var categorySpending = {};
  var totalWeekly = 0;
  
  data.slice(1).forEach(function(row) {
    var date = new Date(row[1]);
    var amount = parseFloat(row[2]);
    var category = row[5];
    var person = row[9];
    var type = row[4];
    
    if (date >= weekAgo && date <= today && type === "Debit" && !isNaN(amount)) {
      if (!weeklySpending[person]) weeklySpending[person] = 0;
      weeklySpending[person] += amount;
      
      if (!categorySpending[category]) categorySpending[category] = 0;
      categorySpending[category] += amount;
      
      totalWeekly += amount;
    }
  });
  
  return {
    weeklySpending: weeklySpending,
    categorySpending: categorySpending,
    totalWeekly: totalWeekly,
    weekStart: Utilities.formatDate(weekAgo, Session.getScriptTimeZone(), "MMM dd"),
    weekEnd: Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM dd")
  };
}

/**
 * Build weekly Telegram message with subscription tracking
 */
function buildWeeklyTelegramMessage(weeklyData, goals, currentSpending, alerts, subscriptions) {
  var message = "📅 *Weekly Summary* (" + weeklyData.weekStart + " - " + weeklyData.weekEnd + ")\n\n";
  
  // Weekly spending by person
  message += "👥 *Weekly Spending*\n";
  for (var person in weeklyData.weeklySpending) {
    message += "👤 " + person + ": ₹" + weeklyData.weeklySpending[person].toFixed(2) + "\n";
  }
  message += "\n📊 *Total Weekly*: ₹" + weeklyData.totalWeekly.toFixed(2) + "\n\n";
  
  // Top spending categories
  message += "📊 *Top Categories This Week*\n";
  var sortedCategories = Object.entries(weeklyData.categorySpending)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 5);
  
  sortedCategories.forEach(function([category, amount]) {
    message += "💳 " + category + ": ₹" + amount.toFixed(2) + "\n";
  });
  message += "\n";
  
  // Monthly goal progress
  if (goals && Object.keys(goals).length > 0) {
    message += "🎯 *Monthly Goal Progress*\n\n";
    var onTrackCount = 0;
    var totalGoals = 0;
    
    for (var category in goals) {
      var goal = goals[category];
      var spent = currentSpending[category] || 0;
      var progress = (spent / goal.target) * 100;
      
      if (progress <= 100) onTrackCount++;
      totalGoals++;
      
      var status = progress < 75 ? "🟢 On Track" : progress < 100 ? "🟡 Warning" : "🔴 Over Budget";
      message += "• " + category + ": " + status + " (₹" + spent.toFixed(0) + "/₹" + goal.target.toFixed(0) + " - " + progress.toFixed(0) + "%)\n";
    }
    
    var overallScore = (onTrackCount / totalGoals) * 100;
    message += "\n🎆 *Overall Score*: " + overallScore.toFixed(0) + "% goals on track\n\n";
  }
  
  // Active Subscriptions Section
  if (subscriptions && Object.keys(subscriptions).length > 0) {
    message += "📱 *Monthly Subscriptions*\n";
    var totalSubscriptions = 0;
    
    for (var service in subscriptions) {
      var amount = subscriptions[service];
      totalSubscriptions += amount;
      message += "🔔 " + service + ": ₹" + amount.toFixed(2) + "\n";
    }
    
    message += "💳 *Total Subscriptions*: ₹" + totalSubscriptions.toFixed(2) + "\n\n";
  }
  
  // Weekly insights and recommendations
  var weeklyInsights = generateWeeklyInsights(weeklyData, goals, currentSpending);
  if (weeklyInsights.length > 0) {
    message += "💡 *Weekly Insights*\n\n";
    weeklyInsights.forEach(function(insight) {
      message += insight + "\n\n";
    });
  }
  
  return message;
}

/**
 * Generate weekly insights and recommendations
 */
function generateWeeklyInsights(weeklyData, goals, currentSpending) {
  var insights = [];
  var today = new Date();
  var dayOfMonth = today.getDate();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var monthProgress = dayOfMonth / daysInMonth;
  
  // Check if spending is on track for the month
  var totalMonthlyTarget = Object.values(goals || {}).reduce((sum, goal) => sum + goal.target, 0);
  var totalCurrentSpending = Object.values(currentSpending || {}).reduce((sum, spent) => sum + spent, 0);
  var expectedSpending = totalMonthlyTarget * monthProgress;
  
  if (totalCurrentSpending < expectedSpending * 0.8) {
    insights.push("🎉 *Excellent*: You're spending 20% less than expected this month!");
  } else if (totalCurrentSpending > expectedSpending * 1.2) {
    insights.push("⚠️ *Alert*: Spending is 20% higher than expected. Consider reducing expenses.");
  }
  
  // Weekly vs daily average insight
  var dailyAverage = weeklyData.totalWeekly / 7;
  if (dailyAverage > 1000) {
    insights.push("📊 *Insight*: Daily average this week: ₹" + dailyAverage.toFixed(0) + ". Consider meal planning to reduce food costs.");
  }
  
  // Category-specific insights
  var topCategory = Object.entries(weeklyData.categorySpending)
    .sort(([,a], [,b]) => b - a)[0];
  
  if (topCategory && topCategory[1] > weeklyData.totalWeekly * 0.4) {
    insights.push("🎯 *Focus Area*: " + topCategory[0] + " accounts for " + 
                 ((topCategory[1] / weeklyData.totalWeekly) * 100).toFixed(0) + 
                 "% of weekly spending. Look for optimization opportunities.");
  }
  
  return insights;
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