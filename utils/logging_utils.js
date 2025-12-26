// NOTE: Still in development, do not use in production.

const LOG_LEVELS = {
  INFO: 'INFO',
  WARN: 'WARN',
  ERROR: 'ERROR',
  CRITICAL: 'CRITICAL'
};

function logToSheet(level, functionName, message, details = '') {
  try {
    const ss = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    const today = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    let logSheet = ss.getSheetByName(today);

    if (!logSheet) {
      logSheet = ss.insertSheet(today, 0);
      const headers = ["Timestamp", "Level", "Function Name", "Message", "Details"];
      logSheet.appendRow(headers).setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "mm/dd/yyyy HH:mm:ss");
    const detailString = (typeof details === 'object') ? JSON.stringify(details, null, 2) : details;
    logSheet.appendRow([timestamp, level, functionName, message, detailString]);
  } catch (e) {
    Logger.log(`[CRITICAL] GAGAL MENULIS LOG: ${e.message}`);
  }
}

const AppLogger = {
  info: (funcName, msg, details) => logToSheet('INFO', funcName, msg, details),
  warn: (funcName, msg, details) => logToSheet('WARN', funcName, msg, details),
  error: (funcName, msg, details) => logToSheet('ERROR', funcName, msg, details),
  critical: (funcName, msg, details) => logToSheet('CRITICAL', funcName, msg, details),
};