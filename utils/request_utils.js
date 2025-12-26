// Cache for department lookups to avoid repeated column searches
const departmentLookupCache = new Map();
const DEPT_CACHE_EXPIRY = 300000; // 5 minutes

function getMultipleColumnValue(sheet, rowValueMap, columnName) {
    const cacheKey = `${getUniqueSheetId(sheet)}_${columnName}`;
    const now = Date.now();

    if (departmentLookupCache.has(cacheKey)) {
        const cached = departmentLookupCache.get(cacheKey);
        if (now - cached.timestamp < DEPT_CACHE_EXPIRY) {
            if (cached.colIndex && rowValueMap) {
                const headers = Object.keys(rowValueMap);
                const header = headers[cached.colIndex - 1];
                const value = rowValueMap[header];
                return { value: value, colIndex: cached.colIndex };
            }
        }
    }

    const result = getColumnValueByKeyword(sheet, rowValueMap, columnName, ACTIVITY_HEADER_ROW_INDEX);

    if (Object.keys(result).length === 0) return {};

    if (result.colIndex) {
        departmentLookupCache.set(cacheKey, { colIndex: result.colIndex, timestamp: now });
    }

    return { value: result.value, colIndex: result.colIndex };
}

// Request Number Generation Utilities
const requestNumberCache = new Map();
const REQUEST_CACHE_EXPIRY_MS = 900000; 
const REQUEST_BATCH_CACHE = new Map(); 
const MAX_CACHE_SIZE = 100;
const pendingSheetUpdates = new Map();

let globalRequestTrackerData = null;
let globalRequestTrackerTimestamp = 0;
const TRACKER_CACHE_EXPIRY = 300000;

function getPersistedRequestCounterMeta(props, prefix) {
    const metaKey = `reqnum.meta.${prefix}`;
    const raw = props.getProperty(metaKey);
    if (!raw) return null;

    try {
        const parsed = JSON.parse(raw);
        if (!Number.isFinite(parsed.rowIndex) || parsed.rowIndex <= 0) {
            props.deleteProperty(metaKey);
            return null;
        }
        if (!Number.isFinite(parsed.colIndex) || parsed.colIndex <= 0) {
            props.deleteProperty(metaKey);
            return null;
        }
        return parsed;
    } catch (e) {
        props.deleteProperty(metaKey);
        return null;
    }
}

function persistRequestCounterMeta(props, prefix, rowIndex, colIndex) {
    if (!Number.isFinite(rowIndex) || rowIndex <= 0) return;
    if (!Number.isFinite(colIndex) || colIndex <= 0) return;

    const metaKey = `reqnum.meta.${prefix}`;
    try {
        props.setProperty(metaKey, JSON.stringify({
            rowIndex,
            colIndex,
            updated: Date.now()
        }));
    } catch (e) {
        Logger.log(`[persistRequestCounterMeta] Failed to persist metadata: ${e.message}`);
    }
}

function getRequestTrackerDataCached() {
    const now = Date.now();
    if (globalRequestTrackerData && (now - globalRequestTrackerTimestamp) < TRACKER_CACHE_EXPIRY) {
        return globalRequestTrackerData;
    }

    const requestSheet = getRequestTracker();
    const lastRow = requestSheet.getLastRow();
    if (lastRow <= 0) {
        globalRequestTrackerData = [];
    } else {
        const countColIndex = getColumnIndex(requestSheet, ColNames.REQUEST_COUNT) || 2;
        const colCount = Math.max(2, countColIndex);
        globalRequestTrackerData = requestSheet.getRange(1, 1, lastRow, colCount).getValues();
    }
    globalRequestTrackerTimestamp = now;
    return globalRequestTrackerData;
}

function queueSheetUpdate(sheet, rowIndex, colIndex, value) {
    const updateKey = `${sheet.getName()}_${rowIndex}_${colIndex}`;
    pendingSheetUpdates.set(updateKey, {
        sheet: sheet,
        rowIndex: rowIndex,
        colIndex: colIndex,
        value: value,
        timestamp: Date.now()
    });
    Utilities.sleep(10);
    flushPendingUpdates();
}

function flushPendingUpdates() {
    if (pendingSheetUpdates.size === 0) return;

    const updates = Array.from(pendingSheetUpdates.values());
    pendingSheetUpdates.clear();

    const updatesBySheet = new Map();
    updates.forEach(update => {
        const sheetName = update.sheet.getName();
        if (!updatesBySheet.has(sheetName)) {
            updatesBySheet.set(sheetName, []);
        }
        updatesBySheet.get(sheetName).push(update);
    });

    updatesBySheet.forEach((sheetUpdates, sheetName) => {
        try {
            sheetUpdates.forEach(update => {
                update.sheet.getRange(update.rowIndex, update.colIndex).setValue(update.value);
            });
        } catch (error) {
            Logger.log(`[flushPendingUpdates] Error updating sheet ${sheetName}: ${error.message}`);
        }
    });
}

function cleanupRequestNumberCache() {
    const now = Date.now();
    const keysToDelete = [];

    for (const [key, value] of requestNumberCache.entries()) {
        if (now - value.timestamp > REQUEST_CACHE_EXPIRY_MS) {
            keysToDelete.push(key);
        }
    }

    keysToDelete.forEach(key => requestNumberCache.delete(key));

    if (requestNumberCache.size > MAX_CACHE_SIZE) {
        const sortedEntries = Array.from(requestNumberCache.entries())
            .sort((a, b) => a[1].timestamp - b[1].timestamp);

        const entriesToRemove = sortedEntries.slice(0, requestNumberCache.size - MAX_CACHE_SIZE);
        entriesToRemove.forEach(([key]) => requestNumberCache.delete(key));
    }
}

function validateCacheEntry(cached, prefix) {
    if (cached == null || typeof cached !== 'object') return false;
    const requiredFields = ['number', 'timestamp', 'rowIndex', 'colIndex'];
    for (const field of requiredFields) {
        if (cached[field] === undefined || cached[field] === null) return false;
    }
    if (typeof cached.number !== 'number' || cached.number < 0) return false;
    return true;
}

function getNextNumber(prefix, isUpdate = true) {
    const start = Date.now();
    const props = PropertiesService.getScriptProperties();
    const propKey = `reqnum.counter.${prefix}`;
    const cacheKey = `${prefix}_number`;
    const metaKey = `reqnum.meta.${prefix}`;
    const persistedMeta = getPersistedRequestCounterMeta(props, prefix);

    try {
        const sheet = getRequestTracker();
        const now = Date.now();
        const cachedEntry = requestNumberCache.get(cacheKey);
        let cacheValid = validateCacheEntry(cachedEntry, prefix) && (now - cachedEntry.timestamp) < REQUEST_CACHE_EXPIRY_MS;

        let rowIndex = persistedMeta ? persistedMeta.rowIndex : -1;
        let colIndex = persistedMeta ? persistedMeta.colIndex : null;
        let cachedNumber = cacheValid ? cachedEntry.number : NaN;
        let shouldPersistMeta = !persistedMeta;

        const validateRowBinding = (candidateRow) => {
            if (!Number.isFinite(candidateRow) || candidateRow <= 0) return false;
            try {
                const currentPrefix = sheet.getRange(candidateRow, 1).getValue();
                return String(currentPrefix) === String(prefix);
            } catch (_) {
                return false;
            }
        };

        if (rowIndex !== -1 && !validateRowBinding(rowIndex)) {
            rowIndex = -1;
            colIndex = null;
            shouldPersistMeta = true;
            props.deleteProperty(metaKey);
        }

        if (rowIndex === -1 && cacheValid) {
            if (validateRowBinding(cachedEntry.rowIndex)) {
                rowIndex = cachedEntry.rowIndex;
                colIndex = colIndex || cachedEntry.colIndex;
                cachedNumber = cachedEntry.number;
            } else {
                cacheValid = false;
                cachedNumber = NaN;
            }
        }

        let data = null;
        if (rowIndex === -1) {
            const lastRow = sheet.getLastRow();
            const firstColRange = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 1) : null;

            if (firstColRange) {
                try {
                    const tf = firstColRange.createTextFinder(String(prefix)).matchEntireCell(true).findNext();
                    if (tf) rowIndex = tf.getRow();
                } catch (_) { }
            }

            if (rowIndex === -1) {
                try { data = getRequestTrackerDataCached(); } catch (_) { data = null; }
                if (Array.isArray(data)) {
                    for (let i = 0; i < data.length; i++) {
                        if (data[i][0] && String(data[i][0]) === String(prefix)) {
                            rowIndex = i + 1;
                            break;
                        }
                    }
                }
            }

            if (rowIndex !== -1) shouldPersistMeta = true;
        }

        if (rowIndex === -1) {
            rowIndex = insertRowValues(sheet, [prefix, 0]);
            shouldPersistMeta = true;
            if (globalRequestTrackerData && Array.isArray(globalRequestTrackerData)) {
                globalRequestTrackerData.push([prefix, 0]);
            }
            if (Array.isArray(data)) data.push([prefix, 0]);
        }

        if (!colIndex) {
            colIndex = getColumnIndex(sheet, ColNames.REQUEST_COUNT) || 2;
            shouldPersistMeta = true;
        }

        if (shouldPersistMeta) {
            persistRequestCounterMeta(props, prefix, rowIndex, colIndex);
        }

        let sheetCur = 0;
        try {
            const cellVal = sheet.getRange(rowIndex, colIndex).getValue();
            const n = parseInt(cellVal, 10);
            sheetCur = Number.isFinite(n) ? n : 0;
        } catch (_) { }

        const propStr = props.getProperty(propKey);
        const propCurParsed = propStr ? parseInt(propStr, 10) : NaN;
        const propCur = Number.isFinite(propCurParsed) ? propCurParsed : NaN;

        let base = sheetCur;
        if (Number.isFinite(propCur)) base = Math.max(base, propCur);
        if (Number.isFinite(cachedNumber)) base = Math.max(base, cachedNumber);

        if (!Number.isFinite(propCur) || propCur < base) {
            props.setProperty(propKey, String(base));
        }

        const newNumber = base + 1;
        props.setProperty(propKey, String(newNumber));

        if (isUpdate) {
            try { sheet.getRange(rowIndex, colIndex).setValue(newNumber); } catch (e) {
                Logger.log(`[getNextNumber] WARN sheet write failed for ${prefix}: ${e.message}`);
            }
            try {
                if (globalRequestTrackerData &&
                    Array.isArray(globalRequestTrackerData) &&
                    globalRequestTrackerData[rowIndex - 1] &&
                    Array.isArray(globalRequestTrackerData[rowIndex - 1])) {
                    globalRequestTrackerData[rowIndex - 1][colIndex - 1] = newNumber;
                }
            } catch (_) { }
        }

        try {
            requestNumberCache.set(cacheKey, {
                number: newNumber,
                timestamp: Date.now(),
                rowIndex, colIndex, version: 1
            });
        } catch (_) { }

        return newNumber;

    } catch (error) {
        Logger.log(`[getNextNumber] ERROR ${prefix}: ${error.message}`);
        return Date.now() % 100000;
    }
}

function getPrefix(sheetNameAbbr, companyName) {
    // Standard format: ABBR/MDM/COMPANY/00001
    return `${sheetNameAbbr}/MDM/${companyName}/`;
}

function generateRequestNumber(sheetName, companyName) {
    const operation = 'generateRequestNumber';
    const priority = 1;
    const maxWaitMs = 300000;
    const t0 = Date.now();

    const sheetNameAbbr = SHEET_ABBR_MAP[sheetName];
    if (!sheetNameAbbr) {
        throw new Error(`[${operation}] Unknown sheetName "${sheetName}" (no abbreviation found).`);
    }

    const rawCompany = (companyName || '').trim();
    // If you need specific handling, ensure input data is normalized before calling this.
    const normalizedCompany = rawCompany;

    const prefix = getPrefix(sheetNameAbbr, normalizedCompany);
    const lockKey = `reqnum:${prefix}`;

    return withKeyLock(lockKey, operation, (_lock, beat) => {
        beat();
        const nextNum = getNextNumber(prefix, /*isUpdate=*/true);
        const reqNum = prefix + String(nextNum).padStart(5, '0');

        Logger.log(
            `[${operation}] Generated ${reqNum} for ${sheetName}/${normalizedCompany} in ${Date.now() - t0}ms`
        );
        return reqNum;
    }, priority, maxWaitMs);
}

function hasActiveRequest(sheet) {
    const [processedDateValues, mdmApprovalDateValues] = getValuesByColumns(
        sheet, [
        ColNames.PROCESSED_DATE,
        ColNames.MDM_APPROVAL_DATE
    ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    const approvalValues = mdmApprovalDateValues.length
        ? mdmApprovalDateValues
        : processedDateValues;

    const emptyIndex = approvalValues.map((val, idx) =>
        !val ? idx : null
    ).filter(idx => idx !== null);

    if (emptyIndex.length === 0) {
        return false;
    }

    const responRequesterValues = getValuesByColumn(
        sheet, ColNames.RESPON_REQUESTER,
        ACTIVITY_HEADER_ROW_INDEX
    );

    const invalidRequesterIndex = responRequesterValues
        .map((val, idx) => (val !== RequesterStatus.COMPLETED || !val ? idx : null))
        .filter(idx => idx !== null);

    const responApproverValues = getValuesByColumn(
        sheet, ColNames.RESPON_APPROVER,
        ACTIVITY_HEADER_ROW_INDEX
    ) || [];

    const invalidApproverIndex = responApproverValues
        .map((val, idx) => (val === ApproverStatus.REJECTED || !val ? idx : null))
        .filter(idx => idx !== null);

    return !emptyIndex.every(idx =>
        invalidRequesterIndex.includes(idx) ||
        invalidApproverIndex.includes(idx)
    );
}

function getRequestClass(sheetName) {
    switch (removeSubmitSuffix(sheetName)) {
        case ActivitySheetNames.PROMOTION:
            RequestClass = RequestPromotion;
            break;
        case ActivitySheetNames.NON_M:
        case ActivitySheetNames.IMAGE:
            RequestClass = RequestWithImage;
            break;
        case ActivitySheetNames.EXTEND_PIR:
            RequestClass = RequestPIRExtend;
            break;
        case ActivitySheetNames.MERCHANDISE:
            RequestClass = RequestMerchandise;
            break;
        case ActivitySheetNames.MASTER_SITE:
            RequestClass = RequestMasterSite;
            break;
        case ActivitySheetNames.MASTER_FINANCE:
            RequestClass = RequestMasterFinance;
            break;
        case ActivitySheetNames.PRICING:
            RequestClass = RequestPricing;
            break;
        case ActivitySheetNames.HIERARCHY:
            RequestClass = RequestHierarchy;
            break;
        case ActivitySheetNames.CUSTOMER:
            RequestClass = RequestCustomer;
            break;
        case ActivitySheetNames.VENDOR:
            RequestClass = RequestVendor;
            break;
        default:
            RequestClass = Request;
            break;
    }

    return RequestClass
}

function shouldExpireRequest(activity, hasRequesterValues, timestamp) {
    if (!hasRequesterValues) {
        return true;
    }

    const { RESPON_REQUESTER } = activity.getActivityValueMap();

    if (RESPON_REQUESTER === RequesterStatus.COMPLETED) {
        let hasPendingApprovals = false;

        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const { prop } = ATTACHMENT_SYNC_CONTEXTS[i];
            const responKey = `RESPON_${prop}`;

            if (activity.hasCol(responKey)) {
                const values = activity.getActivityValueMap();
                const responValue = values[responKey];

                if (!responValue || responValue === '') {
                    hasPendingApprovals = true;
                    break;
                }
            }
        }

        if (hasPendingApprovals) {
            const requestDate = new Date(timestamp);
            const today = new Date();
            const daysDiff = Math.floor((today - requestDate) / (1000 * 60 * 60 * 24));
            const monthlyExpiredLimit = 30; 

            if (daysDiff > monthlyExpiredLimit) {
                return true;
            } else {
                return false;
            }
        }
        return false;
    }
    return true;
}

function getMdmWorkloadEstimates() {
    try {
        const mdmWorkspace = SpreadsheetApp.openById(MDM_WORKSPACE_ID); 
        const taskSheet = mdmWorkspace.getSheetByName("MDM EST");

        if (!taskSheet) return null;

        const data = taskSheet.getDataRange().getValues();
        if (data.length < 2) return {};

        const headers = data[0];
        const mdmWorkloadEstimates = {};
        const totalEstColumnIndex = 1;

        for (let i = 1; i < data.length; i++) { 
            const row = data[i];
            const mdmName = String(row[0] || '').trim().toUpperCase();
            if (!mdmName) continue;

            const estimatesForMdm = {};
            for (let j = 1; j < totalEstColumnIndex; j++) {
                const requestTypeHeader = String(headers[j] || '').trim();
                if (!requestTypeHeader) continue;

                let estimateSeconds = 0;
                const estimateValue = row[j];
                if (typeof estimateValue === 'number' && Number.isInteger(estimateValue)) {
                    estimateSeconds = estimateValue;
                }
                else if (typeof estimateValue === 'string' && estimateValue.includes(':')) {
                    estimateSeconds = parseHms(estimateValue); 
                }
                else {
                    const parsedNum = parseInt(estimateValue, 10);
                    estimateSeconds = !isNaN(parsedNum) ? parsedNum : 0;
                }
                estimatesForMdm[requestTypeHeader] = isFinite(estimateSeconds) ? Math.round(estimateSeconds) : 0;
            }

            let totalEstimateSeconds = 0;
            const totalEstimateValue = row[totalEstColumnIndex];
            if (typeof totalEstimateValue === 'number' && Number.isInteger(totalEstimateValue)) {
                totalEstimateSeconds = totalEstimateValue;
            }
            else if (typeof totalEstimateValue === 'string' && totalEstimateValue.includes(':')) {
                totalEstimateSeconds = parseHms(totalEstimateValue); 
            }
            else {
                const parsedNum = parseInt(totalEstimateValue, 10);
                totalEstimateSeconds = !isNaN(parsedNum) ? parsedNum : 0;
            }
            estimatesForMdm['Total_EST'] = isFinite(totalEstimateSeconds) ? Math.round(totalEstimateSeconds) : 0;

            mdmWorkloadEstimates[mdmName] = estimatesForMdm;
        }

        return mdmWorkloadEstimates;

    } catch (e) {
        Logger.log(`Error reading 'MDM EST' sheet: ${e.message}`);
        return null;
    }
}

function getNextMdmViaRoundRobin(allocationRuleKey, mdmList) {
    if (!mdmList || mdmList.length === 0) return null;

    const cache = CacheService.getScriptCache();
    const counterCacheKey = `scaledRoundRobinCounter_${allocationRuleKey}`;

    let lastIndex = parseInt(cache.get(counterCacheKey) || '-1', 10);
    if (isNaN(lastIndex)) lastIndex = -1;

    const nextIndex = (lastIndex + 1) % mdmList.length;
    cache.put(counterCacheKey, String(nextIndex), 21600); 

    const assignedMdm = mdmList[nextIndex];
    return assignedMdm;
}