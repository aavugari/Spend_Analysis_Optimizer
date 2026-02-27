function extractBankTransactionsWife() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Add headers if missing
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Bank", "Date", "Amount", "Transaction Info", "Transaction Type", "Category", "Card Last 4", "Month", "Year"]);
  }

  // Process last 24 hours only
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var searchDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy/MM/dd");

  extractSBITransactions(sheet, searchDate);
  extractHDFCTransactionsWife(sheet, searchDate);
  extractAmexTransactionsWife(sheet, searchDate);

  formatDateColumn(sheet);
  Logger.log("Wife's transactions extracted successfully!");
}

// Separate functions for full extraction (run individually to avoid timeouts)
function extractSBIOnlyWife() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  extractSBITransactions(sheet, "2025/01/01");
  formatDateColumn(sheet);
}

function extractHDFCOnlyWife() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  extractHDFCTransactionsWife(sheet, "2025/01/01");
  formatDateColumn(sheet);
}

function extractAmexOnlyWife() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  extractAmexTransactionsWife(sheet, "2025/01/01");
  formatDateColumn(sheet);
}

function extractSBITransactions(sheet, searchDate) {
  var threads = GmailApp.search('from:onlinesbicard@sbicard.com subject:"Transaction Alert" after:' + searchDate);
  threads = threads.slice(0, 50);
  var messages = GmailApp.getMessagesForThreads(threads);

  messages.forEach(thread => {
    thread.forEach(message => {
      var body = message.getPlainBody();
      var rawDate = message.getDate();
      var formattedDate = new Date(rawDate);

      var amountMatch = body.match(/Rs\.?\s?([\d,]+\.\d{2})/i);
      var amount = amountMatch ? parseFloat(amountMatch[1].replace(/,/g, '')) : "Not Found";

      var transactionInfo = "Not Found";
      var infoMatch = body.match(/at\s(.+?)\son/i);
      if (infoMatch) {
        transactionInfo = infoMatch[1].trim();
      } else {
        var fallbackMatch = body.match(/spent\s.*?at\s(.*?)(?:\.|\n)/i);
        if (fallbackMatch) {
          transactionInfo = fallbackMatch[1].trim();
        }
      }

      var dateMatch = body.match(/on\s(\d{2}\/\d{2}\/\d{2})/i);
      if (dateMatch) {
        var parts = dateMatch[1].split('/');
        formattedDate = new Date(`20${parts[2]}`, parts[1] - 1, parts[0]);
      }

      var cardMatch = body.match(/ending (\d{4})/i);
      var cardLast4 = cardMatch ? cardMatch[1] : "Unknown";
      
      var transactionType = "Debit";
      var category = getCategoryWife(transactionInfo);
      var month = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "MMMM");
      var year = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "yyyy");

      if (sheet && amount !== "Not Found") {
        var row = ["SBI", formattedDate, amount, transactionInfo, transactionType, category, cardLast4, month, year];
        sheet.appendRow(row);
      }
    });
  });
}

function extractAmexTransactionsWife(sheet, searchDate) {
  var threads = GmailApp.search('from:AmericanExpress@welcome.americanexpress.com OR from:alerts@americanexpress.com after:' + searchDate);
  // No thread limit for Amex to capture all transactions
  var messages = GmailApp.getMessagesForThreads(threads);

  messages.forEach(thread => {
    thread.forEach(message => {
      var body = message.getBody();
      var plainBody = message.getPlainBody();
      var rawDate = message.getDate();
      var formattedDate = new Date(rawDate);

      var dateCutoff = new Date("2025-04-22");
      var amount = "Not Found";
      var transactionInfo = "Not Found";
      var subject = message.getSubject();

      if (formattedDate < dateCutoff && subject.includes("One-Time Password")) {
        // Old format only before 22-April-2025 (OTP emails)
        var amountOldMatch = body.match(/INR\s([\d,]+\.\d{2})/i);
        amount = amountOldMatch ? parseFloat(amountOldMatch[1].replace(/,/g, '')) : "Not Found";

        var infoOldMatch = body.match(/at\s(.*?)\s(?:on|for|is)/i);
        transactionInfo = infoOldMatch ? infoOldMatch[1].replace(/<\/?[^>]+(>|$)/g, "") : "Not Found";
      } else if (plainBody.includes("Merchant:")) {
        // New format after 22-April-2025 (Merchant alerts)
        var merchantMatch = plainBody.match(/Merchant:\s*\n*\s*(.*?)\s*\n/i);
        var amountMatch = plainBody.match(/Amount:\s*\n*INR\s*([\d,]+\.\d{2})/i);
        var dateMatch = plainBody.match(/Date:\s*\n*\s*([0-9]{1,2}\s\w+,\s20\d{2})/i);

        if (merchantMatch) transactionInfo = merchantMatch[1].trim();
        if (amountMatch) amount = parseFloat(amountMatch[1].replace(/,/g, ''));
        if (dateMatch) formattedDate = new Date(dateMatch[1]);
      } else {
        return;
      }

      var cardMatch = plainBody.match(/Account Ending: (\d{5})/) || body.match(/ending in (\d{5})/);
      var cardLast4 = cardMatch ? cardMatch[1] : "Unknown";
      
      var transactionType = "Debit";
      var category = getCategoryWife(transactionInfo);
      var month = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "MMMM");
      var year = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "yyyy");

      if (sheet && amount !== "Not Found") {
        var row = ["Amex", formattedDate, amount, transactionInfo, transactionType, category, cardLast4, month, year];
        sheet.appendRow(row);
      }
    });
  });
}

function extractHDFCTransactionsWife(sheet, searchDate) {
  var threads = GmailApp.search('from:alerts@hdfcbank.net subject:"Alert : Update on your HDFC Bank Credit Card" after:' + searchDate);
  threads = threads.slice(0, 50);
  var messages = GmailApp.getMessagesForThreads(threads);

  messages.forEach(thread => {
    thread.forEach(message => {
      var body = message.getPlainBody();
      var rawDate = message.getDate();
      var formattedDate = new Date(rawDate);

      var amountMatch = body.match(/for\sRs\s([\d,]+\.\d{2})/i);
      var amount = amountMatch ? parseFloat(amountMatch[1].replace(/,/g, '')) : "Not Found";

      var merchantMatch = body.match(/for\sRs\s[\d,]+\.\d{2}\s+at\s+([^\n\r]+?)\s+on/i);
      var transactionInfo = merchantMatch ? merchantMatch[1].trim() : "Not Found";

      var dateMatch = body.match(/on\s(\d{2}-\d{2}-\d{4})/);
      if (dateMatch) {
        var parts = dateMatch[1].split("-");
        formattedDate = new Date(parts[2], parts[1] - 1, parts[0]);
      }

      var cardMatch = body.match(/ending (\d{4})/) || body.match(/\*\*(\d{4})/);
      var cardLast4 = cardMatch ? cardMatch[1] : "Unknown";
      
      var transactionType = "Debit";
      var category = getCategoryWife(transactionInfo);
      var month = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "MMMM");
      var year = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "yyyy");

      if (sheet && amount !== "Not Found") {
        var row = ["HDFC", formattedDate, amount, transactionInfo, transactionType, category, cardLast4, month, year];
        sheet.appendRow(row);
      }
    });
  });
}

function formatDateColumn(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var dateColumn = sheet.getRange(2, 2, lastRow - 1, 1);
    dateColumn.setNumberFormat("MM/dd/yyyy");
  }
}

function getCategoryWife(transactionInfo) {
  const categories = {
    "Rainbow" : "Baby Hospital",
    "Amazon": "Amazon",
    "Swiggy Instamart": "Grocery",
    "Swiggy": "Food",
    "Sri Narsing": "Food",
    "Zomato": "Food",
    "KPN": "Grocery",
    "Zepto": "Grocery",
    "Blinkit": "Grocery",
    "Instamart": "Grocery",
    "Ratnadeep": "Grocery",
    "Super Market": "Grocery",
    "Retail": "Grocery",
    "Apollo": "Health",
    "Health": "Health",
    "Netflix": "Netflix",
    "YouTube": "Youtube Subscription",
    "Google": "Google Subscription",
    "Fuel": "Fuel",
    "HP PAY": "Fuel",
    "Petrol": "Fuel",
    "Donation": "Donation",
    "Milaap": "Donation",
    "Akshaya": "Donation",
    "Dineout": "Dineout",
    "Nykaa": "Shopping",
    "Myntra": "Shopping",
    "BigBasket": "Grocery",
    "LICIOUS": "Food",
    "Food": "Food",
    "Dadus": "Food",
    "Sweet": "Food",
    "Cake": "Food",
    "Voucher": "Amex Voucher",
    "Yashoda": "Health",
    "Vijaya": "Health",
    "Medical": "Health",
    "Hospital": "Health",
    "Drug": "Health",
    "Pista": "Food",
    "Mixture": "Food",
    "Fashion": "Shopping",
    "Car E GH": "Car Maintainance",
    "Automotive": "Car Maintainance",
    "ABR CAFE": "Cafe Niloufer",
    "CAFE NILOUFER": "Cafe Niloufer",
    "Apple": "Subscription",
    "Green Trends": "Grooming",
    "Fresh": "Grocery"
  };

  for (const keyword in categories) {
    if (transactionInfo.toLowerCase().includes(keyword.toLowerCase())) {
      return categories[keyword];
    }
  }
  return "Others";
}