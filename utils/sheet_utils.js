const headerCache = new Map();

// Cache for sheet data to avoid repeated getDataRange() calls
const sheetDataCache = new Map();
const CACHE_EXPIRY_MS = 60000; // 1 minute (reduced from 30 seconds for better performance)

// Performance tracking for optimization analysis
const performanceStats = {
  batchOperations: 0,
  individualOperations: 0,
  cacheHits: 0,
  cacheMisses: 0,
  totalTime: 0,

  reset() {
    this.batchOperations = 0;
    this.individualOperations = 0;
    this.cacheHits = 0;
    this.cacheMisses = 0;
    this.totalTime = 0;
  },

  logSummary() {
    Logger.log(`[Performance Summary] Batch: ${this.batchOperations}, Individual: ${this.individualOperations}, Cache Hits: ${this.cacheHits}, Cache Misses: ${this.cacheMisses}, Total Time: ${this.totalTime}ms`);
  }
};

// Manual cache cleanup function (called when needed)
function cleanupExpiredSheetCache() {
  const now = Date.now();
  let cleanedCount = 0;

  // Clean expired entries from sheet data cache
  for (const [key, value] of sheetDataCache.entries()) {
    if (now - value.timestamp > CACHE_EXPIRY_MS) {
      sheetDataCache.delete(key);
      cleanedCount++;
    }
  }

  // Clean expired entries from header cache periodically
  if (headerCache.size > 100) { // Only clean when cache gets large
    let headerCleanedCount = 0;
    const headerKeys = Array.from(headerCache.keys());
    const maxHeaderCacheSize = 50; // Keep only recent 50 entries

    if (headerKeys.length > maxHeaderCacheSize) {
      const keysToRemove = headerKeys.slice(0, headerKeys.length - maxHeaderCacheSize);
      keysToRemove.forEach(key => {
        headerCache.delete(key);
        headerCleanedCount++;
      });
    }

    if (headerCleanedCount > 0) {
      Logger.log(`[Cache Cleanup] Removed ${headerCleanedCount} old header cache entries`);
    }
  }

  if (cleanedCount > 0) {
    Logger.log(`[Cache Cleanup] Removed ${cleanedCount} expired cache entries`);
  }

  return cleanedCount;
}

function getUniqueSheetId(sheet) {
  const sheetId = sheet.getSheetId();
  const sheetName = sheet.getName();
  return `${sheetId}_${sheetName}`;
}

function getHeaderRowIndex(sheet) {
  const uniqueSheetId = getUniqueSheetId(sheet);
  const cacheKey = `${uniqueSheetId}_headerRowIndex`;

  if (headerCache.has(cacheKey)) {
    return headerCache.get(cacheKey);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i].every(item => item)) {
      let idx = i + 1;
      headerCache.set(cacheKey, idx);
      return idx;
    }
  }

  const defaultHeaderRowIndex = 1;
  headerCache.set(cacheKey, defaultHeaderRowIndex);
  return defaultHeaderRowIndex;
}

function getColumnHeaders(sheet, toUpper = false, headerRowIndex) {
  const uniqueSheetId = getUniqueSheetId(sheet);
  const cacheKey = `${uniqueSheetId}_headers_${toUpper}`;

  if (headerCache.has(cacheKey)) {
    return headerCache.get(cacheKey);
  }

  if (headerRowIndex == null) {
    headerRowIndex = getHeaderRowIndex(sheet);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const rawHeaders = sheet.getRange(headerRowIndex, 1, 1, lastColumn).getValues()[0];

  let effectiveLast = rawHeaders.length;
  while (effectiveLast > 0 && (rawHeaders[effectiveLast - 1] === '' || rawHeaders[effectiveLast - 1] == null)) {
    effectiveLast--;
  }
  const headers = rawHeaders.slice(0, Math.max(effectiveLast, 1));

  const processedHeaders = headers.map(header =>
    (toUpper ? upperAndSeparate(header) : header).toString()
  );

  headerCache.set(cacheKey, processedHeaders);
  return processedHeaders;
}

function getColumnIndex(sheet, colName, headerRowIndex = null) {
  const uniqueSheetId = getUniqueSheetId(sheet);
  const cacheKey = `${uniqueSheetId}_columnIndex_${colName}`;
  if (headerCache.has(cacheKey)) {
    return headerCache.get(cacheKey);
  }
  const headers = getColumnHeaders(sheet, false, headerRowIndex);
  let columnIndex = headers.findIndex(header => header && header == colName) + 1;
  if (columnIndex < 1) return;
  headerCache.set(cacheKey, columnIndex);
  return columnIndex;
}

function handleColParams(sheet, col, headerRowIndex = null) {
  return typeof col === 'number'
    ? col
    : getColumnIndex(sheet, col.toString().trim(), headerRowIndex);
}

function getSheetDataCached(sheet) {
  const uniqueSheetId = getUniqueSheetId(sheet);
  const cacheKey = `${uniqueSheetId}_data`;

  if (sheetDataCache.has(cacheKey)) {
    const cached = sheetDataCache.get(cacheKey);
    if (Date.now() - cached.timestamp < CACHE_EXPIRY_MS) {
      performanceStats.cacheHits++;
      return cached.data;
    }
    sheetDataCache.delete(cacheKey);
  }

  if (sheetDataCache.size > 20) {
    cleanupExpiredSheetCache();
  }

  performanceStats.cacheMisses++;
  const start = Date.now();
  const data = sheet.getDataRange().getValues();
  performanceStats.totalTime += Date.now() - start;

  sheetDataCache.set(cacheKey, {
    data: data,
    timestamp: Date.now()
  });

  return data;
}

function invalidateSheetDataCache(sheet) {
  const uniqueSheetId = getUniqueSheetId(sheet);
  const cacheKey = `${uniqueSheetId}_data`;
  sheetDataCache.delete(cacheKey);
}

function clearRowValuesBetweenColumns(sheet, rowIndex, startColName, endColNameExclusive) {
  if (!sheet || rowIndex < 1) {
    Logger.log('[clearRowValuesBetweenColumns] Invalid sheet or rowIndex provided');
    return false;
  }

  const startColIndex = getColumnIndex(sheet, startColName) + 1;
  if (!startColIndex) {
    Logger.log(`[clearRowValuesBetweenColumns] Start column ${startColName} not found on sheet ${sheet.getName()}`);
    return false;
  }

  let endColIndex;
  if (endColNameExclusive) {
    endColIndex = getColumnIndex(sheet, endColNameExclusive) - 1;
    if (!endColIndex) {
      Logger.log(`[clearRowValuesBetweenColumns] End column ${endColNameExclusive} not found on sheet ${sheet.getName()}`);
      return false;
    }
  } else {
    endColIndex = sheet.getLastColumn();
  }

  const columnCount = endColIndex - startColIndex + 1;
  if (columnCount <= 0) {
    Logger.log(`[clearRowValuesBetweenColumns] No columns to clear between ${startColName} and ${endColNameExclusive || 'end of sheet'} on sheet ${sheet.getName()}`);
    return false;
  }

  sheet.getRange(rowIndex, startColIndex, 1, columnCount).clearContent();
  invalidateSheetDataCache(sheet);
  return true;
}


function getRowIndex(sheet, rowValue, colIndex = 0) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 0) return -1;
    const range = sheet.getRange(1, colIndex + 1, lastRow, 1);
    const finder = range.createTextFinder(String(rowValue)).matchEntireCell(true);
    const match = finder.findNext();
    return match ? match.getRow() : -1;
  } catch (_) {
    const data = getSheetDataCached(sheet);
    for (let i = 0; i < data.length; i++) {
      if (data[i][colIndex] === rowValue) {
        return i + 1;
      }
    }
    return -1;
  }
}

function getRowIndexContain(sheet, rowValue, colIndex = 0) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 0) return -1;
    const range = sheet.getRange(1, colIndex + 1, lastRow, 1);
    const finder = range.createTextFinder(String(rowValue)); // substring match
    const match = finder.findNext();
    return match ? match.getRow() : -1;
  } catch (_) {
    const data = getSheetDataCached(sheet);
    for (let i = 0; i < data.length; i++) {
      if (data[i][colIndex] && data[i][colIndex].toString().includes(rowValue)) {
        return i + 1;
      }
    }
    return -1;
  }
}

function getRowValues(sheet, rowIndex) {
  return sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
}

function getRowValueMap(sheet, rowIndex, toUpper = false, headerRowIndex = null) {
  const headers = getColumnHeaders(sheet, toUpper, headerRowIndex);
  const rowValues = getRowValues(sheet, rowIndex);

  return headers.reduce((map, header, index) => {
    map[header] = rowValues[index] || null;
    return map;
  }, {});
}

function getRowValueMapOptimized(sheet, rowIndex, toUpper = false, headerRowIndex = null) {
  const startTime = new Date();

  const headers = getColumnHeaders(sheet, toUpper, headerRowIndex);

  const numCols = Math.max(headers.length, 1);
  const rowValues = sheet.getRange(rowIndex, 1, 1, numCols).getDisplayValues()[0];

  const result = headers.reduce((map, header, index) => {
    map[header] = rowValues[index] || null;
    return map;
  }, {});

  const elapsed = new Date() - startTime;
  if (elapsed > 100) { 
    Logger.log(`[getRowValueMapOptimized] Completed in ${elapsed}ms for row ${rowIndex}`);
  }

  return result;
}

function getLastNonEmptyRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let row = data.length - 1; row >= 0; row--) {
    if (data[row].some(cell => cell !== "" && cell !== null)) {
      return row + 1; 
    }
  }
  return 0;
}

function getRowValuesMap(sheet, startRow, endRow = null, toUpper = false, headerRowIndex = null) {
  const headers = getColumnHeaders(sheet, toUpper, headerRowIndex);

  if (!endRow) {
    endRow = getLastNonEmptyRow(sheet);
  }

  const rowMaps = [];

  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const rowValues = getRowValues(sheet, rowIndex);

    if (!rowValues || rowValues.length === 0) {
      continue;
    }

    const rowMap = headers.reduce((map, header, index) => {
      const value = rowValues[index] || null;
      map[header] = value;
      return map;
    }, {});

    rowMaps.push(rowMap);
  }

  return rowMaps;  
}

function setValueWithCellRef(sheet, cellRef, value) {
  sheet.getRange(cellRef).setValue(value);
}

function setValueWithIndex(sheet, col, rowIndex, value) {
  const colIndex = handleColParams(sheet, col);
  if (!colIndex) return

  const start = Date.now();

  try {
    sheet.getRange(rowIndex, colIndex).setValue(value);
    invalidateSheetDataCache(sheet);

  } catch (error) {
    Logger.log(`[setValueWithIndex] Error: ${error.message}`);
    throw error;
  }

  const elapsed = Date.now() - start;

  if (elapsed > 1000) { 
    Logger.log(`[Performance] setValueWithIndex took ${elapsed}ms for sheet: ${sheet.getName()}, col: ${colIndex}, row: ${rowIndex}`);
  }

  return { colIndex, rowIndex, value }
}

function insertRowValuesOptimized(sheet, values, insertAt = null, overwrite = true) {
  if (!values || values.length === 0) {
    Logger.log("[insertRowValuesOptimized] No values provided");
    return;
  }

  const start = Date.now();
  const keyValue = values[0];

  let existingRowIndex = getRowIndex(sheet, keyValue);

  if (existingRowIndex !== -1) {
    if (overwrite) {
      const range = sheet.getRange(existingRowIndex, 1, 1, values.length);
      range.setValues([values]);

      const elapsed = Date.now() - start;
      Logger.log(`[insertRowValuesOptimized] Overwrote existing row ${existingRowIndex} in ${elapsed}ms`);
      return existingRowIndex;
    } else {
      const range = sheet.getRange(existingRowIndex, 1, 1, sheet.getLastColumn());
      const existing = range.getValues()[0];

      const merged = existing.map((existingVal, index) => {
        return (values[index] !== null && values[index] !== undefined && values[index] !== '')
          ? values[index]
          : existingVal;
      });

      range.setValues([merged]);

      const elapsed = Date.now() - start;
      Logger.log(`[insertRowValuesOptimized] Merged with existing row ${existingRowIndex} in ${elapsed}ms`);
      return existingRowIndex;
    }
  } else {
    let targetRow;

    if (insertAt && insertAt > 0) {
      sheet.insertRows(insertAt);
      targetRow = insertAt;
    } else {
      targetRow = sheet.getLastRow() + 1;
    }

    const range = sheet.getRange(targetRow, 1, 1, values.length);
    range.setValues([values]);

    invalidateSheetDataCache(sheet);

    const elapsed = Date.now() - start;
    Logger.log(`[insertRowValuesOptimized] Appended new row at ${targetRow} in ${elapsed}ms`);
    return targetRow;
  }
}

function updateRowValues(sheet, rowIndex, updates) {
  if (!updates || updates.length === 0) return;

  const start = Date.now();

  if (updates.length === 1) {
    const { col, value } = updates[0];
    setValueWithIndex(sheet, col, rowIndex, value);
  } else {
    const colIndices = updates.map(u => handleColParams(sheet, u.col)).filter(c => c);
    const values = updates.filter((u, i) => colIndices[i]).map(u => u.value);

    if (colIndices.length > 0) {
      setValuesWithIndexes(sheet, colIndices, rowIndex, values);
    }
  }

  const elapsed = Date.now() - start;
  Logger.log(`[updateRowValues] Updated ${updates.length} values in ${elapsed}ms`);
}

function setValuesWithIndexes(sheet, cols, rowIndex, values) {
  if (!Array.isArray(cols) || !Array.isArray(values) || cols.length !== values.length) {
    throw new Error('Columns and values must be arrays of the same length.');
  }

  const colIndexes = cols.map(col => handleColParams(sheet, col));
  if (colIndexes.some(idx => !idx)) return;

  const start = Date.now();

  const sortedIndexes = [...colIndexes].sort((a, b) => a - b);
  const isContiguous = sortedIndexes.every((col, i) => i === 0 || col === sortedIndexes[i - 1] + 1);

  if (isContiguous && colIndexes.length > 1) {
    const minCol = Math.min(...colIndexes);
    const maxCol = Math.max(...colIndexes);

    const orderedValues = [];
    for (let col = minCol; col <= maxCol; col++) {
      const originalIndex = colIndexes.indexOf(col);
      orderedValues.push(originalIndex !== -1 ? values[originalIndex] : '');
    }

    const range = sheet.getRange(rowIndex, minCol, 1, maxCol - minCol + 1);
    range.setValues([orderedValues]);

    performanceStats.batchOperations++;
    Logger.log(`[setValuesWithIndexes] Batch updated ${colIndexes.length} contiguous columns`);
  } else {
    colIndexes.forEach((colIndex, i) => {
      sheet.getRange(rowIndex, colIndex).setValue(values[i]);
    });

    performanceStats.individualOperations += colIndexes.length;
    Logger.log(`[setValuesWithIndexes] Updated ${colIndexes.length} non-contiguous columns individually`);
  }

  invalidateSheetDataCache(sheet);
  performanceStats.totalTime += Date.now() - start;
  return { colIndexes, rowIndex, values };
}

function insertRowValues(sheet, values, targetRow = null, overwrite = true) {
  const sheetName = sheet.getName();
  const operation = 'insertRowValues';

  const keyValue = values[0];
  if (keyValue === undefined || keyValue === null || keyValue === '') {
    if (targetRow && targetRow > 0) {
      try {
        return withRowLock(sheetName, targetRow, operation, (_rowLock, rowBeat) => {
          const range = sheet.getRange(targetRow, 1, 1, values.length);
          if (overwrite) {
            range.setValues([values]);
            Logger.log(`[${operation}] (no-key) Overwrote row ${targetRow} with values: ${JSON.stringify(values)}`);
          } else {
            const existing = range.getValues()[0];
            const merged = values.map((v, i) => (v !== undefined && v !== null ? v : existing[i]));
            range.setValues([merged]);
            Logger.log(`[${operation}] (no-key) Merged row ${targetRow}. Old: ${JSON.stringify(existing)}, New: ${JSON.stringify(merged)}`);
          }
          invalidateSheetDataCache(sheet);
          rowBeat();
          return targetRow;
        }, 2, 10000);
      } catch (e) {
        Logger.log(`[${operation}] Error (no-key/targetRow=${targetRow}): ${e.message}`);
        throw e;
      }
    }
    throw new Error(`[${operation}] First column (key) must be provided.`);
  }

  return withKeyLock(`rowkey:${sheetName}:${keyValue}`, operation, (_keyLock, keyBeat) => {
    try {
      invalidateSheetDataCache(sheet);

      let rowIndex = targetRow || getRowIndex(sheet, keyValue);

      if (rowIndex === -1) {
        sheet.appendRow(values);
        const newRowIndex = sheet.getLastRow();
        invalidateSheetDataCache(sheet);

        Logger.log(`[${operation}] Appended new row at ${newRowIndex} with values: ${JSON.stringify(values)}`);
        return newRowIndex;
      }

      return withRowLock(sheetName, rowIndex, operation, (rowLock, rowBeat) => {
        const beat = () => { keyBeat(); rowBeat(); };

        beat();
        const range = sheet.getRange(rowIndex, 1, 1, values.length);

        if (overwrite) {
          range.setValues([values]);
          Logger.log(`[${operation}] Overwrote row ${rowIndex} with values: ${JSON.stringify(values)}`);
        } else {
          const existing = range.getValues()[0];
          const merged = values.map((v, i) => (v !== undefined && v !== null ? v : existing[i]));
          range.setValues([merged]);
          Logger.log(`[${operation}] Merged row ${rowIndex}. Old: ${JSON.stringify(existing)}, New: ${JSON.stringify(merged)}`);
        }

        invalidateSheetDataCache(sheet);
        return rowIndex;

      }, 2, 10000);
    } catch (error) {
      Logger.log(`[${operation}] Error: ${error.message}`);
      throw error;
    }
  }, 2, 10000); 
}

function insertRowValuesBatch(sheet, rowsData, operationName = 'insertRowValuesBatch') {
  if (!rowsData || rowsData.length === 0) return [];

  const sheetName = sheet.getName();
  const results = new Array(rowsData.length).fill(null);

  invalidateSheetDataCache(sheet);

  rowsData.forEach((rowData, idx) => {
    const { values, targetRow, overwrite = true } = rowData;
    const keyValue = values && values[0];

    if (keyValue === undefined || keyValue === null || keyValue === '') {
      Logger.log(`[${operationName}] Skipping index ${idx}: missing first-column key.`);
      return;
    }

    withKeyLock(`rowkey:${sheetName}:${keyValue}`, operationName, (keyLock, keyBeat) => {
      try {
        invalidateSheetDataCache(sheet);

        let rowIndex = targetRow || getRowIndex(sheet, keyValue);

        if (rowIndex === -1) {
          sheet.appendRow(values);
          const newRowIndex = sheet.getLastRow();
          results[idx] = newRowIndex;

          invalidateSheetDataCache(sheet);

          Logger.log(`[${operationName}] Appended new row at ${newRowIndex} for key "${keyValue}".`);
          return;
        }

        withRowLock(sheetName, rowIndex, operationName, (rowLock, rowBeat) => {
          const beat = () => { keyBeat(); rowBeat(); };

          beat();

          const range = sheet.getRange(rowIndex, 1, 1, values.length);

          if (overwrite) {
            range.setValues([values]);
            Logger.log(
              `[${operationName}] Overwrote row ${rowIndex} for key "${keyValue}".`
            );
          } else {
            const existing = range.getValues()[0];
            const merged = values.map((v, i) => (v !== undefined && v !== null ? v : existing[i]));
            range.setValues([merged]);
            Logger.log(
              `[${operationName}] Merged row ${rowIndex} for key "${keyValue}".`
            );
          }

          results[idx] = rowIndex;

          invalidateSheetDataCache(sheet);
          beat();

        }, 2, 10000); 

      } catch (e) {
        Logger.log(`[${operationName}] Error processing index ${idx} (key="${keyValue}"): ${e.message}`);
      }
    }, 2, 10000); 
  });

  Logger.log(
    `[${operationName}] Batch processed ${rowsData.length} entries. ` +
    `Created: ${results.filter(r => r !== null && typeof r === 'number').length}, ` +
    `Skipped/Failed: ${results.filter(r => r === null).length}`
  );

  return results;
}

function addSheetEditors(spreadsheet, emails) {
  const fileId = spreadsheet.getId();

  let unique = Array.from(new Set((emails || []).map(e => (e || '').trim()).filter(Boolean)));
  if (unique.length === 0) return;

  try {
    spreadsheet.addEditors(unique);
    Logger.log(`Successfully added ${unique.length} editors to sheet with ID ${fileId}.`);
    return;
  } catch (e) {
    Logger.log(`addSheetEditors batch failed (${e && e.message}). Falling back to chunks.`);
  }

  const chunkSize = 50;
  for (let i = 0; i < unique.length; i += chunkSize) {
    const chunk = unique.slice(i, i + chunkSize);
    try {
      spreadsheet.addEditors(chunk);
      Logger.log(`Added ${chunk.length} editors (chunk) to sheet ${fileId}.`);
      Utilities.sleep(200);
    } catch (e) {
      Logger.log(`addSheetEditors chunk failed (${e && e.message}). Falling back to per-email.`);
      chunk.forEach(email => {
        try {
          Drive.Permissions.insert(
            { role: 'writer', type: 'user', value: email },
            fileId,
            { sendNotificationEmails: 'false' }
          );
        } catch (err) {
          try {
            Drive.Permissions.insert(
              { role: 'writer', type: 'user', value: email },
              fileId,
              { sendNotificationEmails: 'true' }
            );
          } catch (err2) {
            Logger.log(`Failed to add ${email}. Error: ${err2 && err2.message}`);
          }
        }
      });
    }
  }
}

function getValuesByColumn(sheet, col, headerRowIndex = 2) {
  const colIndex = handleColParams(sheet, col, headerRowIndex);
  if (!colIndex) return []

  const lastRow = sheet.getLastRow()
  if (lastRow < headerRowIndex) return [];

  const data = sheet.getRange(headerRowIndex, colIndex, (lastRow - headerRowIndex + 1), 1
  ).getValues();

  return data.map(row => row[0]);
}

function getValuesByColumns(sheet, cols, headerRowIndex = 2) {
  const lastRow = sheet.getLastRow();

  if (lastRow < headerRowIndex) {
    return cols.map(() => []);
  }

  const colIndices = cols.map(col => handleColParams(sheet, col, headerRowIndex));

  const validColData = colIndices.map((colIndex, originalIndex) => ({
    colIndex,
    originalIndex,
    isValid: colIndex !== undefined
  })).filter(item => item.isValid);

  if (validColData.length === 0) {
    return cols.map(() => []);
  }

  const sortedValidCols = [...validColData].sort((a, b) => a.colIndex - b.colIndex);
  const isContiguous = sortedValidCols.length > 1 &&
    sortedValidCols.every((col, i) => i === 0 || col.colIndex === sortedValidCols[i - 1].colIndex + 1);

  let results = new Array(cols.length).fill().map(() => []);

  if (isContiguous && validColData.length > 2) {
    const minCol = Math.min(...validColData.map(v => v.colIndex));
    const maxCol = Math.max(...validColData.map(v => v.colIndex));
    const numCols = maxCol - minCol + 1;

    const batchData = sheet.getRange(headerRowIndex, minCol, lastRow - headerRowIndex + 1, numCols)
      .getDisplayValues();

    validColData.forEach(({ colIndex, originalIndex }) => {
      const batchColIndex = colIndex - minCol;
      results[originalIndex] = batchData.map(row => row[batchColIndex]);
    });

    Logger.log(`[getValuesByColumns] Used batch operation for ${validColData.length} contiguous columns`);
  } else {
    validColData.forEach(({ colIndex, originalIndex }) => {
      results[originalIndex] = sheet.getRange(headerRowIndex, colIndex, lastRow - headerRowIndex + 1, 1)
        .getDisplayValues()
        .map(row => row[0]);
    });

    Logger.log(`[getValuesByColumns] Used individual calls for ${validColData.length} non-contiguous columns`);
  }

  return results;
}

function getValueByColumn(sheet, col, rowIndex, headerRowIndex = 1) {
  const colIndex = handleColParams(sheet, col, headerRowIndex);
  if (!colIndex) return

  return sheet.getRange(rowIndex, colIndex).getValue();
}

function setBackgroundColor(sheet, rowIndex, col, toLastCol = false, color = 'red') {
  const colIndex = handleColParams(sheet, col);
  if (!colIndex) return

  const range = toLastCol
    ? sheet.getRange(rowIndex, colIndex, 1, sheet.getLastColumn() - colIndex + 1)
    : sheet.getRange(rowIndex, colIndex);

  range.setBackground(color);
}

function clearSheetCache(sheet, options = {}) {
  const { skipHeaders = false } = options;
  const uniqueSheetId = getUniqueSheetId(sheet);

  if (!skipHeaders) {
    for (const key of headerCache.keys()) {
      if (key.startsWith(uniqueSheetId)) {
        headerCache.delete(key);
      }
    }
  }

  for (const key of sheetDataCache.keys()) {
    if (key.startsWith(uniqueSheetId)) {
      sheetDataCache.delete(key);
    }
  }

  const logSuffix = skipHeaders ? ' (headers preserved)' : '';
  Logger.log(`Cache cleared for sheet: ${sheet.getName()}${logSuffix}`);
}

function cleanupExpiredCaches() {
  const now = Date.now();

  for (const [key, value] of sheetDataCache.entries()) {
    if (now - value.timestamp > CACHE_EXPIRY_MS) {
      sheetDataCache.delete(key);
    }
  }

  Logger.log(`Cleaned up expired cache entries. Remaining: ${sheetDataCache.size} data cache, ${headerCache.size} header cache`);
}

function getColumnValueByKeyword(sheet, rowValueMap, keyword, headerRowIndex = ACTIVITY_HEADER_ROW_INDEX) {
  const headers = getColumnHeaders(sheet, true, headerRowIndex);

  const headerWithValue = headers.find(header =>
    hasKeywords(header, keyword) && rowValueMap[header]);

  if (!headerWithValue) return {};

  const value = rowValueMap[headerWithValue] || '';
  const headerColIndex = headers.indexOf(headerWithValue) + 1;

  return {
    value: value,
    colIndex: headerColIndex,
    header: headerWithValue
  };
}

function getSelectedRowIndex() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();

  if (!range) {
    Logger.log('No range selected');
    return [];
  }

  var startRow = range.getRow();
  var numRows = range.getNumRows();

  var selectedRows = Array.from({ length: numRows }, (_, i) => startRow + i);

  Logger.log("Selected Row Indexes: " + selectedRows);
  return selectedRows;
}

function getSelectedVisibleRowIndex() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange(); 

  if (!range) {
    Logger.log('No range selected');
    return [];
  }

  var startRow = range.getRow(); 
  var numRows = range.getNumRows(); 
  var selectedRows = Array.from({ length: numRows }, (_, i) => startRow + i);

  var visibleRows = selectedRows.filter(rowIndex => {
    try {
      return !sheet.isRowHiddenByFilter(rowIndex) && !sheet.isRowHiddenByUser(rowIndex);
    } catch (error) {
      Logger.log(`Error checking visibility for row ${rowIndex}: ${error.message}`);
      return true;
    }
  });

  Logger.log("Selected Visible Row Indexes: " + visibleRows);
  Logger.log(`Filtered ${selectedRows.length - visibleRows.length} hidden rows`);
  return visibleRows;
}

function sortRow(enrichedDatas) {
  return enrichedDatas.sort((a, b) => {
    if (a.status === MDMStatus.ON_GOING && b.status !== MDMStatus.ON_GOING) return -1;
    if (a.status !== MDMStatus.ON_GOING && b.status === MDMStatus.ON_GOING) return 1;

    if (!a.status && b.status) return -1; 
    if (a.status && !b.status) return 1;

    return (b.priorityScore || 0) - (a.priorityScore || 0);
  });
}

function prioritySorting(sheet) {
  if (!sheet || sheet.getLastRow() <= 2) { 
    Logger.log(`Sheet '${sheet ? sheet.getName() : 'null'}' empty or invalid, sorting skipped.`);
    return;
  }

  const masterConfig = new MasterConfig();
  const rules = masterConfig.getWeightingRules();
  const timeNow = new Date();
  const dataRange = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const allData = dataRange.getValues();
  const header = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);

  const colIndexes = header.reduce((acc, colName, index) => {
    acc[colName] = index;
    return acc;
  }, {});

  const enrichedData = allData.map((rowData, index) => {
    const priorityScore = masterConfig.getRowScore(rowData, colIndexes, rules, timeNow);

    const processedBy = rowData[colIndexes[ColNames.PROCESSED_BY]];
    const takenDate = rowData[colIndexes[ColNames.TAKEN_DATE]];
    const processStatus = rowData[colIndexes[ColNames.PROCESS_STATUS]];
    const processedDate = rowData[colIndexes[ColNames.PROCESSED_DATE]];

    const isError =
      (!!processedBy && !takenDate) ||
      (!!processStatus && processStatus !== MDMStatus.ON_GOING && !processedDate);

    return {
      rowData: rowData,
      status: String(processStatus || ""),
      priorityScore: priorityScore,
      isError: isError,
      originalIndex: index, 
    };
  });

  const errorRows = enrichedData.filter(item => item.isError);
  const goodRows = enrichedData.filter(item => !item.isError);

  const sortedGoodRows = sortRow(goodRows);

  const finalDataArray = new Array(allData.length);
  errorRows.forEach(item => {
    finalDataArray[item.originalIndex] = item.rowData;
  });

  let goodRowIndex = 0;
  for (let i = 0; i < finalDataArray.length; i++) {
    if (finalDataArray[i] === undefined) {
      if (sortedGoodRows[goodRowIndex]) {
        finalDataArray[i] = sortedGoodRows[goodRowIndex].rowData;
      }
      goodRowIndex++;
    }
  }

  const finalValues = finalDataArray.filter(row => row !== undefined);

  if (finalValues.length > 0) {
    dataRange.clearContent(); 
    sheet.getRange(3, 1, finalValues.length, finalValues[0].length).setValues(finalValues);
    Logger.log(`✅ Sheet '${sheet.getName()}' successfully re-sorted (error rows frozen).`);
  }
}

function logMDMWorkspace(spreadsheet, requestNumber, modification, companyName, department, requestType, oldValue, newValue) {
  try {
    const logSheet = spreadsheet.getSheetByName("MASTER LOG");
    if (!logSheet) {
      Logger.log("ERROR: MASTER LOG sheet not found.");
      return;
    }

    const modifiedBy = Session.getEffectiveUser().getEmail();
    const modifiedTimestamp = getDateNow();

    const logEntry = [
      requestNumber || '',
      modification || '',
      companyName || '',
      department || '',
      requestType || '',
      oldValue || '',
      newValue || '',
      modifiedBy,
      modifiedTimestamp
    ];

    logSheet.appendRow(logEntry);
    Logger.log(`Logged change for Request Number ${requestNumber}: ${modification}`);

  } catch (e) {
    Logger.log(`Failed to log change for Request ${requestNumber}. Error: ${e.message}`);
  }
}

function cleanupCompletedTasks(sheetInput = null) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNamesToProcess = sheetInput ? [sheetInput] : Object.values(MDMSheetNames);

  Logger.log(`[cleanupCompletedTasks] Starting validation, sync, and cleanup process.`);

  sheetNamesToProcess.forEach(sheetName => {
    try {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= ACTIVITY_HEADER_ROW_INDEX) {
        return;
      }

      // (Optional) This mutates ordering; safe to run without a sheet lock.
      prioritySorting(sheet);

      const headers = sheet.getRange(
        ACTIVITY_HEADER_ROW_INDEX, 1, 1, sheet.getLastColumn()
      ).getValues()[0];

      const colIndexes = {
        processStatus: headers.indexOf(ColNames.PROCESS_STATUS),
        takenDate: headers.indexOf(ColNames.TAKEN_DATE),
        estimatedTimeFinished: headers.indexOf(ColNames.ESTIMATED_TIME_FINISHED),
        processedDate: headers.indexOf(ColNames.PROCESSED_DATE),
        feedbackStatus: headers.indexOf(ColNames.FEEDBACK_STATUS),
        attachment: headers.indexOf(ColNames.ATTACHMENT),
      };

      if (Object.values(colIndexes).some(index => index === -1)) {
        Logger.log(`[cleanupCompletedTasks] ERROR: Important columns not found in sheet ${sheetName}.`);
        return;
      }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const values = sheet.getRange(
        ACTIVITY_HEADER_ROW_INDEX + 1, 1,
        lastRow - ACTIVITY_HEADER_ROW_INDEX,
        lastCol
      ).getValues();

      const rowsToProcess = [];

      values.forEach((row, idx) => {
        const processStatus = row[colIndexes.processStatus];
        const takenDate = row[colIndexes.takenDate];
        const estTimeFinished = row[colIndexes.estimatedTimeFinished];
        const processedDate = row[colIndexes.processedDate];
        const feedbackStatus = row[colIndexes.feedbackStatus];
        const attachment = row[colIndexes.attachment];

        const isFullyPopulated = takenDate && processStatus && estTimeFinished && processedDate && feedbackStatus;
        if (!isFullyPopulated) return;

        let shouldProcess = false;
        if (processStatus === MDMStatus.SEND_BACK) {
          if (attachment === 'NO ATTACHMENT') {
            shouldProcess = true;
          }
        } else if (processStatus !== MDMStatus.ON_GOING) {
          shouldProcess = true;
        }

        if (shouldProcess) {
          rowsToProcess.push(idx + ACTIVITY_HEADER_ROW_INDEX + 1); // convert to absolute row index
        }
      });

      if (rowsToProcess.length === 0) return;

      Logger.log(`[cleanupCompletedTasks] Found ${rowsToProcess.length} rows to sync and delete in sheet ${sheetName}.`);

      // Delete from bottom to avoid index shifting
      rowsToProcess.sort((a, b) => b - a).forEach(rowIndex => {
        try {
          withRowLock(sheetName, rowIndex, 'cleanupCompletedTasks', (rowLock, rowBeat) => {
            const beat = () => rowBeat();

            if (rowIndex > sheet.getLastRow()) {
              Logger.log(`[cleanupCompletedTasks] Skipping row ${rowIndex} (no longer exists).`);
              return;
            }

            const RequestClass = getRequestClass(sheet.getName());
            const request = new RequestClass(sheet, rowIndex);
            beat();
            request.activityHandler.copyDataToMaster();
            beat();

            sheet.deleteRow(rowIndex);
            beat();

            invalidateSheetDataCache(sheet);

            Logger.log(`[cleanupCompletedTasks] ✔️ Row ${rowIndex} synced & deleted from ${sheetName}.`);
          }, 3, 5000);
        } catch (e) {
          Logger.log(`[cleanupCompletedTasks] ERROR processing row ${rowIndex}: ${e.message}`);
        }
      });

    } catch (e) {
      Logger.log(`[cleanupCompletedTasks] ERROR processing sheet ${sheetName}: ${e.message}`);
    }
  });

  Logger.log('[cleanupCompletedTasks] Cleanup process completed.');
}



function runMasterSheetAuditAndFix() {
  Logger.log("Starting 4-in-1 Audit...");

  const file1_MasterSS = SpreadsheetApp.getActiveSpreadsheet();
  const file2_ChildSS = SpreadsheetApp.openById(MDM_WORKSPACE_ID); 
  const HEADER_ROW_AUDIT = ACTIVITY_HEADER_ROW_INDEX; 
  
  const sheetNamesToAudit = [
      "STATUS/LISTING", "BASIC DATA", "NON M", "MERCHANDISE", 
      "EXTEND PIR", "PROMOTION", "SOURCE LIST", "MASTER DATA", "IMAGE", 
      "HIERARCHY", "BOM", "MASTER FINANCE", "PRICING", "MASTER SITE", 
      "PROFIT CENTER", "VENDOR"
  ];

  let missingInChildRequests = [];
  let duplicateRequests = [];
  let unallocatedRequests = [];
  let missingTotalTaskRequests = [];
  try {
    
    missingInChildRequests = findMissingInChildSheets(file1_MasterSS, file2_ChildSS, sheetNamesToAudit, HEADER_ROW_AUDIT);
    if (missingInChildRequests.length > 0) {
        const missingReqNums = missingInChildRequests.map(item => item.reqNum).join(', ');
        Logger.log(`Audit 1 Complete: Found ${missingInChildRequests.length} missing requests in child sheets: ${missingReqNums}`);
    } else {
        Logger.log(`Audit 1 Complete: Found ${missingInChildRequests.length} missing requests in child sheets.`);
    }
    const errorsFromDupeSheet = findErrorsFromDupeSheet(file1_MasterSS);
    duplicateRequests = errorsFromDupeSheet.duplicateRequests;
    unallocatedRequests = errorsFromDupeSheet.unallocatedRequests;
    missingTotalTaskRequests = errorsFromDupeSheet.missingTotalTaskRequests;
    
    if (duplicateRequests.length > 0) {
        const duplicateReqNums = duplicateRequests.map(item => `${item.reqNum} (in ${item.sheetName})`).join('; ');
        Logger.log(`Audit 2 Complete: Found ${duplicateRequests.length} duplicate requests from "DUPE" sheet: ${duplicateReqNums}`);
    } else {
        Logger.log(`Audit 2 Complete: Found ${duplicateRequests.length} duplicate requests from "DUPE" sheet.`);
    }

    if (unallocatedRequests.length > 0) {
        const unallocatedReqNums = unallocatedRequests.map(item => `${item.reqNum} (in ${item.sheetName})`).join(', ');
        Logger.log(`Audit 3 Complete: Found ${unallocatedRequests.length} unallocated requests (Approved but Empty Processed By) from "DUPE" sheet: ${unallocatedReqNums}`);
    } else {
        Logger.log(`Audit 3 Complete: Found ${unallocatedRequests.length} unallocated requests from "DUPE" sheet.`);
    }
    
    if (missingTotalTaskRequests.length > 0) {
        const missingTaskReqNums = missingTotalTaskRequests.map(item => `${item.reqNum} (in ${item.sheetName})`).join(', ');
        Logger.log(`Audit 4 Complete: Found ${missingTotalTaskRequests.length} requests (Empty Total Task but Completed) from "DUPE" sheet: ${missingTaskReqNums}`);
    } else {
        Logger.log(`Audit 4 Complete: Found ${missingTotalTaskRequests.length} requests from "DUPE" sheet.`);
    }

  } catch (e) {
    Logger.log(`ERROR during audit: ${e.message}`);
  }
  
  if (missingInChildRequests.length > 0) {
    Logger.log(`--- Starting Fix 1: copyDataToChild for ${missingInChildRequests.length} requests... ---`);
    fixMissingInChild(file1_MasterSS, missingInChildRequests);
  }

  if (duplicateRequests.length > 0) {
    Logger.log(`--- Starting Fix 2: fixOnSubmit for ${duplicateRequests.length} duplicate requests... ---`);
    fixDuplicateRequests(file1_MasterSS, duplicateRequests);
  }

  if (unallocatedRequests.length > 0) {
    Logger.log(`--- Starting Fix 3: handleRequestApproved for ${unallocatedRequests.length} requests... ---`);
    fixUnallocatedRequests(file1_MasterSS, unallocatedRequests);
  }

  if (missingTotalTaskRequests.length > 0) {
    Logger.log(`--- Starting Fix 4: handleTotalTask for ${missingTotalTaskRequests.length} requests... ---`);
    fixMissingTotalTask(file1_MasterSS, missingTotalTaskRequests);
  }
  
  Logger.log("Audit 4-in-1 Complete.");
}

function findErrorsFromDupeSheet(file1_MasterSS) {
    Logger.log("[Audit 2,3,4] Reading from 'DUPE' sheet...");
    const duplicateRequests = [];
    const unallocatedRequests = [];
    const missingTotalTaskRequests = [];

    try {
        const dupeSheet = file1_MasterSS.getSheetByName("DUPE");
        if (!dupeSheet) {
            throw new Error("'DUPE' sheet not found.");
        }

        const dataStartRow = 3; 
        if (dupeSheet.getLastRow() < dataStartRow) {
            Logger.log("[Audit 2,3,4] 'DUPE' sheet is empty (no data from row 3).");
            return { duplicateRequests, unallocatedRequests, missingTotalTaskRequests };
        }

        const dataRange = dupeSheet.getRange(dataStartRow, 1, dupeSheet.getLastRow() - (dataStartRow - 1), 10);
        const values = dataRange.getValues();

        values.forEach((row, index) => {
            const rowIndex = index + dataStartRow;

            // Audit 2: Duplicate (Col A & B)
            const dupReqNum = row[0]; 
            const dupSheetName = row[1]; 
            if (dupReqNum && dupSheetName) {
                duplicateRequests.push({ reqNum: dupReqNum, sheetName: dupSheetName, rowIndex: null }); 
            }

            // Audit 3: Unallocated (Col C & D)
            const unallocatedSheetName = row[2]; 
            const unallocatedReqNum = row[3]; 
            if (unallocatedReqNum && unallocatedSheetName) {
                unallocatedRequests.push({ reqNum: unallocatedReqNum, sheetName: unallocatedSheetName, rowIndex: null }); 
            }

            // Audit 4: Missing Total Task (Col I & J)
            const missingTaskSheetName = row[8]; 
            const missingTaskReqNum = row[9]; 
            if (missingTaskReqNum && missingTaskSheetName) {
                missingTotalTaskRequests.push({ reqNum: missingTaskReqNum, sheetName: missingTaskSheetName, rowIndex: null }); 
            }
        });

    } catch (e) {
        Logger.log(`[Audit 2,3,4] FAILED to read "DUPE" sheet: ${e.message}`);
    }
    
    return {
        duplicateRequests: duplicateRequests,
        unallocatedRequests: unallocatedRequests,
        missingTotalTaskRequests: missingTotalTaskRequests
    };
}

function loadChildSheetRequestNumbers(file2_ChildSS, HEADER_ROW_AUDIT) {
  const childSheetCache = new Map();
  const all_sheets_file_2 = file2_ChildSS.getSheets();
  Logger.log(`[loadChildSheetRequestNumbers] Starting load. Found ${all_sheets_file_2.length} total sheets in File 2.`);
  all_sheets_file_2.forEach(sheet => {
    const sheetName = sheet.getName();
    try {
      if (!Object.values(MDMSheetNames).includes(sheetName.toUpperCase())) { 
        return; 
      }
      Logger.log(`[loadChildSheetRequestNumbers] Processing sheet: ${sheetName}...`);
      const values = sheet.getDataRange().getValues();
      if (values.length < HEADER_ROW_AUDIT) {
        Logger.log(`[loadChildSheetRequestNumbers] Skipping sheet: ${sheetName} (Fewer rows than HEADER_ROW_AUDIT).`);
        return;
      }
      const header = values[HEADER_ROW_AUDIT - 1]; 
      const data = values.slice(HEADER_ROW_AUDIT); 
      const reqNumColIdx = header.indexOf("Request Number");
      if (reqNumColIdx === -1) {
        Logger.log(`[loadChildSheetRequestNumbers] Skipping sheet: ${sheetName} ('Request Number' column not found).`);
        return;
      }
      const requestNumberSet = new Set();
      data.forEach(row => {
        const reqNum = row[reqNumColIdx];
        if (reqNum) requestNumberSet.add(String(reqNum).trim());
      });
      childSheetCache.set(sheetName.trim().toUpperCase(), requestNumberSet);
      const requestNumbersList = Array.from(requestNumberSet).join(', ');
      Logger.log(`[loadChildSheetRequestNumbers] SUCCESS: ${sheetName} loaded with ${requestNumberSet.size} unique request numbers.`);
    } catch (e) {
      Logger.log(`[loadChildSheetRequestNumbers] FAILED to load child sheet ${sheetName}: ${e.message}`);
    }
  });
  Logger.log(`[loadChildSheetRequestNumbers] Finished: Loaded ${childSheetCache.size} child sheets to cache.`);
  return childSheetCache;
}

function findMissingInChildSheets(file1_MasterSS, file2_ChildSS, sheetNamesToAudit, HEADER_ROW_AUDIT) {
  const childCache = loadChildSheetRequestNumbers(file2_ChildSS, HEADER_ROW_AUDIT);
  const missingRequests = [];

  sheetNamesToAudit.forEach(sheetName => {
    try {
      const ws = file1_MasterSS.getSheetByName(sheetName);
      if (!ws) throw new Error(`Sheet ${sheetName} not found.`);
      
      const values = ws.getDataRange().getValues();
      if (values.length < HEADER_ROW_AUDIT) return;
      
      const header = values[HEADER_ROW_AUDIT - 1];
      const data = values.slice(HEADER_ROW_AUDIT);
      
      const reqNumIdx = header.indexOf("Request Number");
      const procByIdx = header.indexOf("Processed By");
      const procDateIdx = header.indexOf("Processed Date");

      if (reqNumIdx === -1 || procByIdx === -1 || procDateIdx === -1) {
          Logger.log(`Audit 1: Skipping sheet '${sheetName}' (incomplete columns).`);
          return;
      }

      data.forEach((row, i) => {
        const request_number = row[reqNumIdx];
        const processed_by = row[procByIdx];
        const process_status = row[procDateIdx]; 

        if (!request_number || !processed_by) return;
        
        if (!process_status || String(process_status).trim() === "") {
          const processed_by_upper = String(processed_by).trim().toUpperCase();
          const df_target = childCache.get(processed_by_upper);

          if (!df_target) {
            Logger.log(`Audit 1: Sheet '${processed_by_upper}' (from ${request_number}) not found in File 2 cache.`);
            return;
          }

          if (!df_target.has(String(request_number))) {
            missingRequests.push({ reqNum: request_number, sheetName: sheetName });
          }
        }
      });
    } catch (e) {
      Logger.log(`Audit 1: Failed processing sheet '${sheetName}': ${e.message}`);
    }
  });
  return missingRequests;
}

// =================================================================
// FIX FUNCTIONS
// =================================================================

function fixMissingInChild(file1_MasterSS, requests) {
  requests.forEach(item => {
    try {
      const masterSheet = file1_MasterSS.getSheetByName(item.sheetName);
      if (!masterSheet) throw new Error(`Sheet ${item.sheetName} not found.`);
      
      const rowIndex = getRowIndex(masterSheet, item.reqNum); 
      if (rowIndex === -1) throw new Error(`Request ${item.reqNum} not found anymore.`);

      const RequestClass = getRequestClass(item.sheetName); 
      const request = new RequestClass(masterSheet, rowIndex); 
      const result = request.activityHandler.copyDataToChild(); 

      if (result.success) {
        Logger.log(`FIX 1 SUCCESS: ${item.reqNum} copied to child sheet.`);
      } else {
        Logger.log(`FIX 1 FAILED: ${item.reqNum}. Message: ${result.message}`);
      }
    } catch (e) {
      Logger.log(`FIX 1 ERROR: Failed to process ${item.reqNum}. Error: ${e.message}`);
    }
  });
}

function fixDuplicateRequests(file1_MasterSS, requests) {
  requests.forEach(item => {
    try {
      const masterSheet = file1_MasterSS.getSheetByName(item.sheetName);
      if (!masterSheet) throw new Error(`Sheet ${item.sheetName} not found.`);
      
      const rowIndex = getRowIndex(masterSheet, item.reqNum); 
      if (rowIndex === -1) {
        Logger.log(`FIX 2: Could not find row for duplicate ${item.reqNum} in sheet ${item.sheetName}. Might be fixed already.`);
        return;
      }

      const RequestClass = getRequestClass(item.sheetName); 
      const request = new RequestClass(masterSheet, rowIndex); 
      
      const oldReqNum = item.reqNum;
      const { ATTACHMENT } = request.activity.getActivityValueMap(false); 
      const oldAttachmentFile = ATTACHMENT ? DriveApp.getFileById(extractSheetId(ATTACHMENT)) : null; 
      
      request.activity.updateRequestNumber(null); 
      request.handleOnSubmit(); 
      
      const { REQUEST_NUMBER: newReqNum } = request.activity.getActivityValueMap(true); 
      
      if (oldAttachmentFile && newReqNum && newReqNum !== oldReqNum) {
        const oldAttachmentName = oldAttachmentFile.getName();
        const newAttachmentName = oldAttachmentName.replace(oldReqNum, newReqNum);
        oldAttachmentFile.rename(newAttachmentName);
        Logger.log(`FIX 2: Old attachment ${oldAttachmentName} renamed to ${newAttachmentName}`);
      }
      
      Logger.log(`FIX 2 SUCCESS: ${oldReqNum} replaced with ${newReqNum} in sheet ${item.sheetName}.`);

    } catch (e) {
      Logger.log(`FIX 2 ERROR: Failed to process duplicate ${item.reqNum}. Error: ${e.message}`);
    }
  });
}

function fixUnallocatedRequests(file1_MasterSS, requests) {
  requests.forEach(item => {
    try {
      const masterSheet = file1_MasterSS.getSheetByName(item.sheetName);
      if (!masterSheet) throw new Error(`Sheet ${item.sheetName} not found.`);
      
      const rowIndex = getRowIndex(masterSheet, item.reqNum); 
      if (rowIndex === -1) {
         Logger.log(`FIX 3: Could not find row for ${item.reqNum} in sheet ${item.sheetName}. Might be fixed already.`);
        return;
      }
      
      const RequestClass = getRequestClass(item.sheetName); 
      const request = new RequestClass(masterSheet, rowIndex); 
      
      const result = request.requestHandler.handleRequestApproved(); 
      
      if (result) {
        Logger.log(`FIX 3 SUCCESS: ${item.reqNum} allocated successfully.`);
      } else {
        Logger.log(`FIX 3 FAILED: ${item.reqNum}. Allocation failed.`);
      }
    } catch (e) {
      Logger.log(`FIX 3 ERROR: Failed to process ${item.reqNum}. Error: ${e.message}`);
    }
  });
}

function fixMissingTotalTask(file1_MasterSS, requests) {
  requests.forEach(item => {
    try {
      const masterSheet = file1_MasterSS.getSheetByName(item.sheetName);
      if (!masterSheet) throw new Error(`Sheet ${item.sheetName} not found.`);
      
      const rowIndex = getRowIndex(masterSheet, item.reqNum); 
      if (rowIndex === -1) {
        Logger.log(`FIX 4: Could not find row for ${item.reqNum} in sheet ${item.sheetName}. Might be fixed already.`);
        return;
      }
      
      const RequestClass = getRequestClass(item.sheetName); 
      const request = new RequestClass(masterSheet, item.rowIndex); 
      
      const result = request.requestHandler.handleTotalTask(); 
      
      if (result) {
        Logger.log(`FIX 4 SUCCESS: ${item.reqNum} Total Task recalculated successfully.`);
      } else {
        Logger.log(`FIX 4 FAILED: ${item.reqNum}. Failed to calculate Total Task.`);
      }
    } catch (e) {
      Logger.log(`FIX 4 ERROR: Failed to process ${item.reqNum}. Error: ${e.message}`);
    }
  });
}