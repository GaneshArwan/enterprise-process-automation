class MasterConfig {
    constructor() {
        this.spreadsheet = SpreadsheetApp.openById(MASTER_CONFIGURATION_UID);
        this.cache = new Map();
        this.cacheTimestamp = new Map();
        this.CACHE_DURATION = 5 * 60 * 1000; // 5 minutes in milliseconds
    }

    /**
     * Generic method to get cached data or fetch fresh data if cache is expired
     * @param {string} sheetName - Name of the sheet to fetch data from
     * @param {string[]} columns - Array of column names to fetch
     * @param {number} headerRowIndex - Row index where headers are located
     * @returns {Object} Object with column data arrays
     */
    getCachedSheetData(sheetName, columns, headerRowIndex = 1) {
        const cacheKey = `${sheetName}_${columns.join('_')}`;
        const now = Date.now();
        
        // Check if we have valid cached data
        if (this.cache.has(cacheKey) && 
            this.cacheTimestamp.has(cacheKey) && 
            (now - this.cacheTimestamp.get(cacheKey)) < this.CACHE_DURATION) {
            return this.cache.get(cacheKey);
        }

        // Fetch fresh data
        const sheet = this.spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.error(`[MasterConfig] Failed to get "${sheetName}" sheet`);
            return null;
        }

        const columnData = getValuesByColumns(sheet, columns, headerRowIndex);
        const result = {};
        columns.forEach((col, index) => {
            result[col] = columnData[index] || [];
        });

        // Cache the result
        this.cache.set(cacheKey, result);
        this.cacheTimestamp.set(cacheKey, now);

        return result;
    }

    /**
     * Clear all cached data (useful for testing or forcing refresh)
     */
    clearCache() {
        this.cache.clear();
        this.cacheTimestamp.clear();
    }

    /**
     * Clear specific sheet cache
     * @param {string} sheetName - Name of sheet to clear from cache
     */
    clearSheetCache(sheetName) {
        const keysToDelete = Array.from(this.cache.keys()).filter(key => key.startsWith(sheetName));
        keysToDelete.forEach(key => {
            this.cache.delete(key);
            this.cacheTimestamp.delete(key);
        });
    }

    getApproverList(context) {
        const {
            companyCode,
            department,
            requestType,
            level,
            useDefault = true
        } = context;

        // Use cached data retrieval
        const data = this.getCachedSheetData('Approvers', [
            'Company Code - Name',
            'Department', 
            'Request Type',
            'Email Address',
            'Level Order',
            'Active'
        ]);

        if (!data) return [];

        const { 
            'Company Code - Name': companies,
            'Department': depts,
            'Request Type': reqs,
            'Email Address': emails,
            'Level Order': levels,
            'Active': actives
        } = data;

        const ALL = ApproverConfiguration.ALL;

        // Build lookup map with optimized indexing
        const lookup = new Map();
        for (let i = 1; i < companies.length; i++) {
            // Pre-validate row to skip early
            if (String(companies[i]).trim() !== companyCode ||
                String(levels[i]).trim() !== String(level) ||
                String(actives[i]).toUpperCase() !== 'TRUE') {
                continue;
            }

            const deptKey = String(depts[i]).trim();
            const reqKey = String(reqs[i]).trim();
            const key = `${deptKey}|${reqKey}|${level}`;
            
            if (!lookup.has(key)) {
                lookup.set(key, []);
            }

            // Process emails once and add to list
            const emailList = String(emails[i])
                .split('\n')
                .map(e => e.trim())
                .filter(Boolean);
            
            lookup.get(key).push(...emailList);
        }

        // Define search keys in priority order
        const keysToTry = [
            `${department}|${requestType}|${level}`,
            ...(useDefault ? [
                `${ALL}|${requestType}|${level}`,
                `${department}|${ALL}|${level}`,
                `${ALL}|${ALL}|${level}`
            ] : [])
        ];

        // Find first non-empty match
        for (const key of keysToTry) {
            const emailsForKey = lookup.get(key);
            if (emailsForKey && emailsForKey.length) {
                const normalized = emailsForKey.map(e => e.trim());
                // Check for NO_APPROVER
                if (normalized.length === 0 || normalized.includes(NO_APPROVER)) {
                    return [];
                }
                return normalized;
            }
        }

        return [];
    }

    /**
     * Returns the baseline seconds for a given requestType and totalTask,
     * or null if none found / on error.
     *
     * @param {{ requestType: string, totalTask: string|number }} context
     * @returns {number|null}
     */
    getBaseline({ requestType, totalTask }) {
        // 1) Validate totalTask
        const total = Number(totalTask);
        if (isNaN(total)) {
            console.error(
                `[ConfigurationMaster] Invalid totalTask: "${totalTask}"`
            );
            return null;
        }

        // 2) Use cached data retrieval
        const data = this.getCachedSheetData('Baseline', [
            'REQUEST TYPE', 
            'TASK RANGE', 
            'SECONDS',
            'IS TASK BASELINE',
        ]);

        if (!data) {
            console.error('[ConfigurationMaster] Failed to get "Baseline" sheet data');
            return null;
        }

        const {
            'REQUEST TYPE': types,
            'TASK RANGE': ranges,
            'SECONDS': secs,
            'IS TASK BASELINE': isTaskBaseline,
        } = data;

        // 3) Scan rows 1…n looking for a match
        for (let i = 1; i < types.length; i++) {
            if (String(types[i]).trim() !== requestType) continue;

            const rangeStr = String(ranges[i]).trim();
            // single regex for "n-m" or "n+"
            const m = rangeStr.match(/^(\d+)(?:-(\d+)|\+)?$/);
            if (!m) {
                console.warn(
                    `[ConfigurationMaster] Skipping bad TASK RANGE: "${rangeStr}"`
                );
                continue;
            }

            const min = parseInt(m[1], 10);
            const max = m[2]
                ? parseInt(m[2], 10)
                : rangeStr.endsWith('+')
                    ? Infinity
                    : min;
            
            // Lakukan konversi yang aman ke boolean
            const isBaselineTaskBoolean = String(isTaskBaseline[i]).trim().toLowerCase() === 'true';

            // if totalTask falls in [min…max], return the SECONDS value
            if (total >= min && total <= max) {
                return { 
                    baseline: parseFloat(secs[i]), 
                    isTaskBaseline: isBaselineTaskBoolean
                };
            }
        }

        console.warn(
            `[ConfigurationMaster] No baseline for "${requestType}" with totalTask=${total}`
        );
        return { baseline: null, isTaskBaseline: null };;
    }

    getWorkAllocation(context) {
        const {
            companyCode,
            requestType,
            department,
            useDefault = true
        } = context;

        const ALL = 'ALL';

        // Use cached data retrieval
        const data = this.getCachedSheetData('Work Allocation', [
            'COMPANY CODE',
            'ACTIVITIES',
            'DEPT',
            'PIC',
            'BACKUP 1',
            'BACKUP 2',
            'BACKUP 3',
            'DEFAULT'
        ]);

        if (!data) {
            console.error('[ConfigurationMaster] Failed to get "Work Allocation" sheet data');
            return null;
        }

        const {
            'COMPANY CODE': companyCodes,
            'ACTIVITIES': activitiesArr,
            'DEPT': depts,
            'PIC': pics,
            'BACKUP 1': backup1Arr,
            'BACKUP 2': backup2Arr,
            'BACKUP 3': backup3Arr,
            'DEFAULT': defaultsArr
        } = data;

        // 1) build a Map<"company|activity|dept", { pic, backups, default }>
        const lookup = new Map();
        for (let i = 1; i < companyCodes.length; i++) {
            const cc = String(companyCodes[i]).trim();
            const act = String(activitiesArr[i]).trim();
            const dept = String(depts[i]).trim();

            // skip incomplete rows
            if (!cc || !act || !dept) continue;

            const key = `${cc}|${act}|${dept}`;
            const pic = String(pics[i]).trim();
            const backups = [backup1Arr[i], backup2Arr[i], backup3Arr[i]]
                .map(b => String(b).trim())
                .filter(Boolean);
            const defPic = String(defaultsArr[i]).trim();

            lookup.set(key, { 
                pic,
                backups,
                default: defPic
            });
        }

        // 2) define the keys-to-try in priority order
        const keysToTry = [
            // exact match
            `${companyCode}|${requestType}|${department}`,
            // fallbacks (only if useDefault)
            ...(useDefault ? [
                // wildcard activity
                `${companyCode}|${ALL}|${department}`,
                // wildcard dept
                `${companyCode}|${requestType}|${ALL}`,
                // both wildcards
                `${companyCode}|${ALL}|${ALL}`,
            ] : [])
        ];

        // 3) pick the first one that exists in our Map
        for (const key of keysToTry) {
            const alloc = lookup.get(key);
            if (alloc) {
                return alloc;
            }
        }

        // 4) no match
        console.warn(
            `[ConfigurationMaster] No work allocation for company="${companyCode}", ` +
            `activity="${requestType}", dept="${department}"`
        );
        return null;
    }

    getWeightingRules() {
        // Use cached data retrieval for weighting rules
        const data = this.getCachedSheetData('Priority Weight', [
            'COLUMN', 'OPERATOR', 'VALUE1', 'VALUE2', 'WEIGHT', 'IMPORTANCE'
        ], 1);

        if (!data) {
            console.error('[ConfigurationMaster] Failed to get "Priority Weight" sheet data');
            return [];
        }

        // Convert to array of rule objects
        const rules = [];
        const {
            'COLUMN': columns,
            'OPERATOR': operators,
            'VALUE1': values1,
            'VALUE2': values2,
            'WEIGHT': weights,
            'IMPORTANCE': importance
        } = data;

        for (let i = 1; i < columns.length; i++) {
            if (columns[i]) { // Skip empty rows
                rules.push({
                    COLUMN: columns[i],
                    OPERATOR: operators[i],
                    VALUE1: values1[i],
                    VALUE2: values2[i],
                    WEIGHT: weights[i],
                    IMPORTANCE: importance[i]
                });
            }
        }

        return rules;
    }

    getRowScore(rowData, colIndexes, rules, timeNow) {
        let totalScore = 0;
        let uniqueRuleColumns = [...new Set(rules.map(r => r.COLUMN))];
        if (rowData[colIndexes[ColNames.VALID_FROM]]) {
            uniqueRuleColumns = uniqueRuleColumns.filter(col => col !== ColNames.TIMESTAMP_ENTRY && col !== ColNames.REQUEST_TYPE);
        }

        for (const columnName of uniqueRuleColumns) {
            if (colIndexes[columnName] === undefined) continue;
            
            const value = rowData[colIndexes[columnName]];
            const relevantRules = rules.filter(r => r.COLUMN === columnName);

            const foundRule = relevantRules.find(r => {
                const op = r.OPERATOR, v1 = r.VALUE1, v2 = r.VALUE2;
                switch (columnName) {
                    case ColNames.TIMESTAMP_ENTRY:
                        if (!(value instanceof Date)) return false;
                        const hoursDiff = Math.floor((timeNow - value) / 3600000);
                        if (op === 'gt') return hoursDiff > v1;
                        if (op === 'eq') return hoursDiff == v1;
                        if (op === 'lte') return hoursDiff <= v1;
                        return false;

                    case ColNames.VALID_FROM:
                        if (!(value instanceof Date)) return false;
                        const today = new Date(); today.setHours(0, 0, 0, 0);
                        const valueDate = new Date(value.getTime()); valueDate.setHours(0, 0, 0, 0);
                        const diffDays = Math.floor((valueDate - today) / 86400000);
                        if (op === 'lte') return diffDays <= v1;
                        if (op === 'eq') return diffDays == v1;
                        if (op === 'gt') return diffDays > v1;
                        return false;

                    default:
                        if (op === 'in') return String(v1).split('\n').map(s => s.trim()).includes(String(value));
                        if (!isNaN(value) && value !== "") {
                            if (op === 'lte') return value <= v1;
                            if (op === 'gt') return value > v1;
                            if (op === 'gt_lte') return value > v1 && value <= v2;
                        }
                        return false;
                }
            });

            if (foundRule) {
                totalScore += parseFloat(foundRule.WEIGHT) * parseFloat(foundRule.IMPORTANCE);
            }

            if (totalScore >= 1) {
                totalScore = 1; 
                break;
            }
        }

        return Math.round(totalScore * 100) / 100;
    }

    /**
     * Get configuration data with optimized caching
     * @param {string} configType - Type of config ('Request', 'Attachment', etc.)
     * @param {string} key - Optional key to look up specific row
     * @returns {Object|Sheet} Configuration data or sheet
     */
    getConfigData(configType, key = null) {
        if (!key) {
            // Return the sheet if no specific key requested
            return this.spreadsheet.getSheetByName(configType);
        }

        // For specific lookups, use caching
        const cacheKey = `config_${configType}_${key}`;
        const now = Date.now();
        
        if (this.cache.has(cacheKey) && 
            this.cacheTimestamp.has(cacheKey) && 
            (now - this.cacheTimestamp.get(cacheKey)) < this.CACHE_DURATION) {
            return this.cache.get(cacheKey);
        }

        const sheet = this.spreadsheet.getSheetByName(configType);
        if (!sheet) {
            console.error(`[MasterConfig] Failed to get "${configType}" sheet`);
            return null;
        }

        let result;
        if (configType === 'Request') {
            const configRowIndex = getRowIndexContain(sheet, key);
            console.log(`[MasterConfig] getRowIndexContain returned row index: ${configRowIndex} for key "${key}"`);
            
            // Check if row index is valid (greater than 0)
            if (configRowIndex <= 0) {
                console.warn(`[MasterConfig] Invalid row index (${configRowIndex}) for key "${key}" in ${configType} sheet`);
                return null;
            }
            
            // Additional validation: check if the sheet has enough rows
            const lastRow = sheet.getLastRow();
            console.log(`[MasterConfig] Sheet "${configType}" has ${lastRow} rows`);
            
            if (configRowIndex > lastRow) {
                console.warn(`[MasterConfig] Row index ${configRowIndex} exceeds sheet last row (${lastRow}) for key "${key}"`);
                return null;
            }
            
            // Check if the sheet has any data rows (beyond headers)
            if (lastRow <= 1) {
                console.warn(`[MasterConfig] Sheet "${configType}" has no data rows (only ${lastRow} row(s))`);
                return null;
            }
            
            try {
                result = getRowValueMap(sheet, configRowIndex);
            } catch (error) {
                console.error(`[MasterConfig] Error getting row values for key "${key}" at row ${configRowIndex}:`, error);
                return null;
            }
        } else if (configType === 'Attachment') {
            const rowIndex = getRowIndex(sheet, key);
            if (rowIndex <= 0) {
                console.warn(`[MasterConfig] Invalid row index (${rowIndex}) for key "${key}" in ${configType} sheet`);
                return null;
            }
            result = sheet.getRange(rowIndex, AttachmentValues.UID_COL_INDEX).getValue();
        } else {
            // Generic row lookup
            const rowIndex = getRowIndex(sheet, key);
            if (rowIndex <= 0) {
                console.warn(`[MasterConfig] Invalid row index (${rowIndex}) for key "${key}" in ${configType} sheet`);
                return null;
            }
            result = getRowValueMap(sheet, rowIndex);
        }

        // Cache the result
        this.cache.set(cacheKey, result);
        this.cacheTimestamp.set(cacheKey, now);

        return result;
    }

    /**
     * Batch multiple config operations for better performance
     * @param {Object[]} operations - Array of {type, params} objects
     * @returns {Object[]} Array of results corresponding to operations
     */
    batchConfigOperations(operations) {
        const results = [];
        
        for (const operation of operations) {
            try {
                let result;
                switch (operation.type) {
                    case 'baseline':
                        result = this.getBaseline(operation.params);
                        break;
                    case 'approvers':
                        result = this.getApproverList(operation.params);
                        break;
                    case 'workAllocation':
                        result = this.getWorkAllocation(operation.params);
                        break;
                    case 'config':
                        result = this.getConfigData(operation.params.configType, operation.params.key);
                        break;
                    default:
                        console.warn(`[MasterConfig] Unknown operation type: ${operation.type}`);
                        result = null;
                }
                results.push(result);
            } catch (error) {
                console.error(`[MasterConfig] Error in batch operation ${operation.type}:`, error);
                results.push(null);
            }
        }
        
        return results;
    }

    /**
     * Force refresh of all cached data (useful for testing or data updates)
     */
    forceRefresh() {
        this.clearCache();
        console.log('[MasterConfig] All cached data cleared - fresh data will be fetched on next requests');
    }
}


// Shared instance for better performance
let _sharedMasterConfig = null;

function getSharedMasterConfig() {
    if (!_sharedMasterConfig) {
        _sharedMasterConfig = new MasterConfig();
    }
    return _sharedMasterConfig;
}

function getConfigSpreadsheet(sheetName = null) {
    const masterConfig = getSharedMasterConfig();
    
    if (!sheetName) {
        return masterConfig.spreadsheet.getSheets()[0];
    }

    return masterConfig.spreadsheet.getSheetByName(sheetName);
}

function getConfigApproverNew(companyCode, department, requestType, level, useDefault = true) {
    const sheet = getConfigSpreadsheet('Approvers');
    if (!sheet) {
        return [];  // Return an empty array if the sheet is not found
    }

    const ALL = ApproverConfiguration.ALL;         // your "DEFAULT"/ALL key

    const headerRowIndex = 1;
    const [
        companies, depts, reqs,
        emails, levels, actives
    ] = getValuesByColumns(
        sheet,
        [
            'Company Code - Name',
            'Department',
            'Request Type',
            'Email Address',
            'Level Order',
            'Active'
        ],
        headerRowIndex
    );

    /**
     * Returns all emails matching (companyCode, deptKey, reqKey, level, TRUE)
     * Splits emails by '\n', trims them, and filters out empty values
     */
    function fetchFor(deptKey, reqKey) {
        const out = [];
        for (let i = 1; i < companies.length; i++) {
            if (
                String(companies[i]).trim() === companyCode &&
                String(depts[i]).trim() === String(deptKey).trim() &&
                String(reqs[i]).trim() === String(reqKey).trim() &&
                String(levels[i]) === String(level) &&
                String(actives[i]).toUpperCase() === 'TRUE'
            ) {
                // Split the emails by '\n', trim each email, and filter out empty ones
                const processedEmails = emails[i]
                    .split('\n')
                    .map(email => email.trim())
                    .filter(Boolean);

                out.push(...processedEmails);  // Add the processed emails to the result
            }
        }
        return out;
    }

    // 1) dept & requestType
    let result = fetchFor(department, requestType);

    // 2) ALL & requestType
    if (result.length === 0 && useDefault) {
        result = fetchFor(ALL, requestType);
    }

    // 3) dept & ALL
    if (result.length === 0 && useDefault) {
        result = fetchFor(department, ALL);
    }

    // 4) ALL & ALL
    if (result.length === 0 && useDefault) {
        result = fetchFor(ALL, ALL);
    }

    // Normalize and check for NO_APPROVER
    const normalizedResult = result.map(email => email.trim());  // Trim all emails in the result
    if (normalizedResult.length === 0 || normalizedResult.includes(NO_APPROVER)) {
        return [];
    }

    return result;  // Return the array of approvers
}

function getConfigApprover(companyName, department, requestType, useDefault = true) {
    const sheet = SpreadsheetApp.openById(APPROVER_CONFIGURATION_UID).getSheetByName(companyName);
    if (!sheet) {
        return undefined;
    }

    const getValue = (col, row) => {
        return row !== -1 ? getValueByColumn(sheet, col, row) : null;
    };

    const checkApprover = (col, row) => {
        const value = getValue(col, row);
        if (String(value) === NO_APPROVER) {
            return undefined;
        }

        const approvers = isNotEmpty(value)
            ? value.split('\n').map(e => e.trim()).filter(Boolean)
            : null;
        return approvers;
    };

    const departmentRow = getRowIndex(sheet, department);
    const defaultRow = getRowIndex(sheet, ApproverConfiguration.DEFAULT);

    let approver;

    approver = checkApprover(requestType, departmentRow);
    if (approver !== null || !useDefault) {
        return approver;
    }

    approver = checkApprover(requestType, defaultRow);
    if (approver !== null) {
        return approver;
    }

    approver = checkApprover(ApproverConfiguration.DEFAULT, departmentRow);
    if (approver !== null) {
        return approver;
    }

    approver = checkApprover(ApproverConfiguration.DEFAULT, defaultRow);
    if (approver !== null) {
        return approver;
    }

    return undefined;
}


function getRequestConfig(companyName = null) {
    const masterConfig = getSharedMasterConfig();
    
    if (!companyName) {
        return masterConfig.getConfigData('Request');
    }

    // Try to get config with the original company name
    let config = masterConfig.getConfigData('Request', companyName);
    
    // If no config found, try with 'FALLBACK' as lookup key
    if (!config) {
        console.log(`[getRequestConfig] No configuration found for "${companyName}", trying FALLBACK`);
        config = masterConfig.getConfigData('Request', 'FALLBACK');
        
        if (!config) {
            console.error(`[getRequestConfig] CRITICAL: No configuration found for "${companyName}" or FALLBACK. Please check your configuration sheet.`);
            console.error(`[getRequestConfig] Make sure you have a row with "${companyName}" or "FALLBACK" in the Request configuration sheet.`);
        } else {
            console.log(`[getRequestConfig] Successfully found FALLBACK configuration for "${companyName}"`);
        }
    }
    
    return config;
}

function getAttachmentUID(key) {
    const masterConfig = getSharedMasterConfig();
    return masterConfig.getConfigData('Attachment', key);
}

function getArchiveFolder() {
    return DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
}

function getChildSpreadsheet(Activity) {
    const activityValueMap = Activity.getActivityValueMap(true,true);
    let sheetName = Activity.sheet.getName();

    //List of target sheet names that should be used as is
    let lookupKey = Activity.getCompanyName();
    sheetName = activityValueMap.PROCESSED_BY;

    console.log("Lookup Key: " + lookupKey);

    const config = getRequestConfig(lookupKey);
    
    if (!config || !config[CHILD_SPREADSHEET_KEY]) {
        throw new Error(`[getChildSpreadsheet] No configuration found for "${lookupKey}" or FALLBACK`);
    }

    const childUrl = config[CHILD_SPREADSHEET_KEY];
    Logger.log(`[getChildSpreadsheet] Opening Child Spreadsheet: ${childUrl}`);

    return SpreadsheetApp.openByUrl(childUrl).getSheetByName(sheetName);
}


function getMasterSpreadsheet(sheetName = null) {
    const masterUID = getAttachmentUID('Master');
    const masterSpreadsheet = SpreadsheetApp.openById(masterUID)

    if (!sheetName) return masterSpreadsheet;

    const finalSheet = masterSpreadsheet.getSheetByName(sheetName);
    return finalSheet
}

function getRequestTracker() {
    return getMasterSpreadsheet(SheetNames.REQUEST);
    // return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetNames.REQUEST);
}

/**
 * Force refresh of master config cache (useful for testing or after config updates)
 */
function refreshMasterConfigCache() {
    if (_sharedMasterConfig) {
        _sharedMasterConfig.forceRefresh();
    }
}

/**
 * Get cache statistics for monitoring performance
 * @returns {Object} Cache statistics
 */
function getMasterConfigCacheStats() {
    if (!_sharedMasterConfig) {
        return { cacheSize: 0, message: 'No shared config instance' };
    }
    
    const now = Date.now();
    const cacheEntries = Array.from(_sharedMasterConfig.cache.keys());
    const validEntries = cacheEntries.filter(key => {
        const timestamp = _sharedMasterConfig.cacheTimestamp.get(key);
        return timestamp && (now - timestamp) < _sharedMasterConfig.CACHE_DURATION;
    });
    
    return {
        totalEntries: cacheEntries.length,
        validEntries: validEntries.length,
        expiredEntries: cacheEntries.length - validEntries.length,
        cacheDuration: _sharedMasterConfig.CACHE_DURATION,
        cacheKeys: validEntries
    };
}

/**
 * Optimized config retrieval for high-frequency operations
 * @param {string} configType - Type of configuration
 * @param {Object} params - Parameters for the operation
 * @returns {*} Configuration result
 */
function getOptimizedConfig(configType, params) {
    const masterConfig = getSharedMasterConfig();
    
    switch (configType) {
        case 'baseline':
            return masterConfig.getBaseline(params);
        case 'approvers':
            return masterConfig.getApproverList(params);
        case 'workAllocation':
            return masterConfig.getWorkAllocation(params);
        case 'weightingRules':
            return masterConfig.getWeightingRules();
        default:
            console.warn(`[getOptimizedConfig] Unknown config type: ${configType}`);
            return null;
    }
}

function getMdmDistributionMatrix() {
    const cache = CacheService.getScriptCache();
    const CACHE_KEY = 'MDM_DISTRIBUTION_MATRIX_CACHE'; // Key cache diperbarui agar sesuai
    const CACHE_DURATION = 21600; // 6 jam

    // 1. Coba ambil dari cache
    try {
        const cachedData = cache.get(CACHE_KEY);
        if (cachedData) {
            Logger.log("[getMdmDistributionMatrix] Cache HIT. Menggunakan data matrix dari cache.");
            return JSON.parse(cachedData);
        }
    } catch (e) {
        Logger.log(`[getMdmDistributionMatrix] Gagal membaca cache: ${e.message}. Mengambil data baru.`);
    }

    // 2. Cache MISS: Baca dari sheet
    Logger.log("[getMdmDistributionMatrix] Cache MISS. Membaca sheet 'Distribution'...");

    try {
        const configSS = SpreadsheetApp.openById(MASTER_CONFIGURATION_UID); //
        // --- PERUBAHAN NAMA SHEET DI SINI ---
        const matrixSheet = configSS.getSheetByName("Distribution"); 
        // ------------------------------------

        if (!matrixSheet) {
            Logger.log("Error: Sheet 'Distribution' tidak ditemukan.");
            return null;
        }

        const data = matrixSheet.getDataRange().getValues();
        if (data.length < 2) return {};

        // Baris 1 adalah Header Nama MDM (Mulai dari Kolom B / Index 1)
        const mdmHeaders = data[0].slice(1).map(name => String(name || '').trim().toUpperCase()); 
        
        const matrix = {};

        // Loop baris data (Mulai Baris 2)
        for (let i = 1; i < data.length; i++) {
            const requestType = String(data[i][0] || '').trim(); // Kolom A: Request Type
            if (!requestType) continue;

            const eligibleMdms = [];
            
            for (let j = 0; j < mdmHeaders.length; j++) {
                const mdmName = mdmHeaders[j];
                if (!mdmName) continue;

                const isChecked = data[i][j + 1]; 

                if (isChecked === true || String(isChecked).toUpperCase() === 'TRUE') {
                    eligibleMdms.push(mdmName);
                }
            }

            if (eligibleMdms.length > 0) {
                matrix[requestType] = eligibleMdms;
            }
        }

        // 3. Simpan ke Cache
        try {
            cache.put(CACHE_KEY, JSON.stringify(matrix), CACHE_DURATION);
            Logger.log(`[getMdmDistributionMatrix] Matrix disimpan di cache selama ${CACHE_DURATION} detik.`);
        } catch (e) {
            Logger.log(`[getMdmDistributionMatrix] Gagal menyimpan ke cache: ${e.message}`);
        }

        return matrix;

    } catch (e) {
        Logger.log(`Error reading 'Distribution': ${e.message}`);
        return null;
    }
}

function clearMdmDistributionCache() {
    CacheService.getScriptCache().remove('MDM_DISTRIBUTION_MATRIX_CACHE');
    Logger.log("Cache MDM Distribution berhasil dihapus.");
}