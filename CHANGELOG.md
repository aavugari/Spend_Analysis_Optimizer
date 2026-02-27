# Changelog

## [v2.0.1] - 2025-01-28

### Removed
- **Mobile Dashboard**: Temporarily removed PWA dashboard implementation for later development
- **API Components**: Removed Google Apps Script Web App API and related files
- **PWA Files**: Removed manifest.json, service worker, and mobile interface

## [v2.0.0] - 2025-01-28

### Major Refactoring & Performance Optimization

#### 🚀 New Features
- **Bank-wise Extraction Functions**: Separate functions for each bank to avoid timeouts
- **Enhanced Telegram Notifications**: Comprehensive budget tracking and family spending breakdown
- **Amex Format Logic**: Dual format support with April cutoff dates (Husband: April 1, Wife: April 22)
- **24-Hour Processing**: Incremental data collection to avoid refreshing existing data
- **Backup Functionality**: Automated Google Drive backups with error handling

#### ⚡ Performance Improvements
- **Thread Limit Optimization**: Reduced limits for ICICI/HDFC (50 threads), unlimited for Amex
- **Timeout Prevention**: Individual bank extraction functions to stay within 6-minute execution limits
- **Memory Optimization**: Efficient message processing and data handling

#### 🔧 Technical Enhancements
- **Error Handling**: Comprehensive validation for Telegram API and Drive operations
- **Code Organization**: Separated Telegram functionality from transaction extraction
- **Data Reliability**: Enhanced parsing logic for all bank formats
- **Category Mapping**: Improved transaction categorization for both users

#### 📁 File Structure
- `Transactions-Husband.js` - Husband transaction extractor (ICICI, HDFC, Amex)
- `Transactions-Wife.js` - Wife transaction extractor (SBI, HDFC, Amex)  
- `Transaction-Merger.js` - Enhanced merger with Telegram notifications and backup
- `REFACTORING_GUIDE.md` - Technical documentation and implementation details

#### 🐛 Bug Fixes
- Fixed Amex transaction parsing for both old and new email formats
- Resolved Telegram API empty message errors
- Fixed Drive API file operation errors
- Corrected syntax errors and duplicate function issues

#### 💰 Budget Tracking
- Monthly budget limits: Food (₹15,000), Grocery (₹8,000), Shopping (₹5,000)
- Real-time spending alerts and weekly summaries
- Family spending breakdown and category-wise analysis

---

## [v1.0.0] - Initial Version
- Basic Gmail transaction extraction
- Simple Google Sheets integration
- Manual processing workflow