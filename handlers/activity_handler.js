class ActivityHandler {
    constructor(activity) {
        this.activity = activity;
        this.sheet = activity.sheet;
        this.rowIndex = activity.rowIndex;
    }

    /**
     * Copies data from the current activity to a target sheet.
     * 
     * Row Existence Behavior:
     * - Checks if a row with the same key (first column, typically REQUEST_NUMBER) already exists
     * - If row exists and overwrite=true: replaces the entire row with new data
     * - If row exists and overwrite=false: merges new data with existing row (only updates non-empty values)
     * - If row doesn't exist: creates a new row regardless of overwrite setting
     * 
     * Overwrite Settings by Method:
     * - copyDataToMaster(): overwrite=false (merge with existing data)
     * - copyDataToChild(): overwrite=true (replace existing data completely)
     * - copyDataSubmissionToMaster(): overwrite=true (replace existing data completely)
     * 
     * @param {Sheet} targetSheet - The sheet to copy data to
     * @param {number} headerRowIndex - The row index containing headers (default: ACTIVITY_HEADER_ROW_INDEX)
     * @param {boolean} overwrite - Whether to overwrite existing rows completely (true) or merge (false)
     * @param {boolean} forceRefresh - Whether to force refresh the activity value map
     * @returns {Object} Result object containing success status, target row index, and operation details
     */
    copyData(
        targetSheet,
        headerRowIndex = ACTIVITY_HEADER_ROW_INDEX,
        overwrite = true,
        forceRefresh = false
    ) {
        const operationName = `copyData_${this.activity.constructor.name}`;
        const startTime = new Date();
        Logger.log(`[${operationName}] Starting copy operation`);

        try {
            // Use parallel operations where possible
            const promises = [];

            // Step 1: Get activity value map (cached)
            const step1Start = new Date();
            const activityValueMap = this.activity.getActivityValueMap(true, forceRefresh);
            const step1Time = new Date() - step1Start;
            Logger.log(`[${operationName}] getActivityValueMap completed in ${step1Time}ms`);

            // Step 2: Get headers (cached)
            const step2Start = new Date();
            const headers = getColumnHeaders(
                targetSheet, true,
                headerRowIndex
            );
            const step2Time = new Date() - step2Start;
            Logger.log(`[${operationName}] getColumnHeaders completed in ${step2Time}ms`);

            // Step 3: Map values
            const step3Start = new Date();
            const rowValues = headers.map(header => activityValueMap[header]);
            const step3Time = new Date() - step3Start;
            Logger.log(`[${operationName}] Value mapping completed in ${step3Time}ms`);

            // Enhanced validation for meaningful data
            // Check if we have at least a key value (first column, usually REQUEST_NUMBER)
            const primaryKeyValue = rowValues[0];
            const hasNonEmptyValues = rowValues.some(val => val !== null && val !== undefined && val !== '');

            if (!primaryKeyValue) {
                Logger.log(`[${operationName}] Warning: No key value (first column) found. Skipping operation.`);
                return {
                    targetSheet: targetSheet,
                    targetRowIndex: null,
                    rowValues: rowValues,
                    success: false,
                    message: "No key value to copy"
                };
            }

            if (!hasNonEmptyValues) {
                Logger.log(`[${operationName}] Warning: No meaningful data to copy (all values are empty). Key: ${primaryKeyValue}`);
                return {
                    targetSheet: targetSheet,
                    targetRowIndex: null,
                    rowValues: rowValues,
                    success: false,
                    message: "No meaningful data to copy"
                };
            }

            Logger.log(`[${operationName}] Data validation passed - Key: ${primaryKeyValue}, Non-empty values: ${rowValues.filter(v => v).length}/${rowValues.length}`);

            // Step 4: Fast row existence check using cached data
            const step4Start = new Date();
            const keyValue = primaryKeyValue; // Use the validated key value
            let existingRowIndex = -1;

            // Use cached sheet data for faster lookup
            const sheetId = getUniqueSheetId(targetSheet);
            if (sheetDataCache.has(`${sheetId}_data`)) {
                const cachedData = sheetDataCache.get(`${sheetId}_data`);
                if (Date.now() - cachedData.timestamp < CACHE_EXPIRY_MS) {
                    // Search in cached data
                    for (let i = 0; i < cachedData.data.length; i++) {
                        if (cachedData.data[i][0] === keyValue) {
                            existingRowIndex = i + 1;
                            break;
                        }
                    }
                } else {
                    existingRowIndex = getRowIndex(targetSheet, keyValue);
                }
            } else {
                existingRowIndex = getRowIndex(targetSheet, keyValue);
            }

            const rowExists = existingRowIndex !== -1;
            const step4Time = new Date() - step4Start;
            Logger.log(`[${operationName}] Row existence check completed in ${step4Time}ms`);

            let actionTaken;
            if (rowExists) {
                actionTaken = overwrite ? "updated existing row" : "merged with existing row";
                Logger.log(`[${operationName}] Row exists at index ${existingRowIndex}. Will ${actionTaken}.`);
            } else {
                actionTaken = "created new row";
                Logger.log(`[${operationName}] Row does not exist. Will ${actionTaken}.`);
            }

            // Step 5: Optimized insert/update
            const step5Start = new Date();
            const targetRowIndex = insertRowValues(targetSheet, rowValues, null, overwrite);
            const step5Time = new Date() - step5Start;
            Logger.log(`[${operationName}] insertRowValues completed in ${step5Time}ms`);

            const totalTime = new Date() - startTime;
            Logger.log(`[${operationName}] Successfully ${actionTaken} at row ${targetRowIndex} in sheet "${targetSheet.getName()}" - Total time: ${totalTime}ms`);
            Logger.log(`[${operationName}] Breakdown: getValueMap=${step1Time}ms, getHeaders=${step2Time}ms, mapping=${step3Time}ms, rowCheck=${step4Time}ms, insert=${step5Time}ms`);

            return {
                targetSheet: targetSheet,
                targetRowIndex: targetRowIndex,
                rowValues: rowValues,
                success: true,
                message: `Data copied successfully - ${actionTaken}`,
                rowExists: rowExists,
                actionTaken: actionTaken
            };

        } catch (error) {
            const totalTime = new Date() - startTime;
            Logger.log(`[${operationName}] Error copying data after ${totalTime}ms: ${error.message}`);

            // Return error information for better debugging
            return {
                targetSheet: targetSheet,
                targetRowIndex: null,
                rowValues: null,
                success: false,
                message: `Copy failed: ${error.message}`,
                error: error
            };
        }
    }

    copyDataSubmissionToMaster() {
        const operationName = 'copyDataSubmissionToMaster';

        try {
            const sheetName = this.activity.sheet.getName();
            const baseName = this.activity.getBaseName();

            if (sheetName === baseName) {
                Logger.log(`[${operationName}] Skipping copy data to master`);
                return;
            }

            const masterSheet = getMasterSpreadsheet(baseName);

            Logger.log(`[${operationName}] Copying submission data to master sheet "${masterSheet.getName()}"`);

            // Panggil copyData dan kirim sinyal 'forceRefresh = true'
            // Set overwrite = true for submissions to ensure complete data replacement
            const result = this.copyData(masterSheet, ACTIVITY_HEADER_ROW_INDEX, true, true);

            if (result.success) {
                Logger.log(`[${operationName}] Successfully copied submission to master at row ${result.targetRowIndex} - ${result.actionTaken}`);
            } else {
                Logger.log(`[${operationName}] Failed to copy submission: ${result.message}`);
            }

            return result;
        } catch (error) {
            Logger.log(`[${operationName}] Error in copyDataSubmissionToMaster: ${error.message}`);
            throw error;
        }
    }

    // Copy data from to the child spreadsheet
    copyDataToChild() {
        const operationName = 'copyDataToChild';

        // try {
        //     const childSheet = getChildSpreadsheet(this.activity);
        //     Logger.log(`[${operationName}] Copying data to child sheet "${childSheet.getName()}"`);

        //     // Use overwrite = true to update existing rows completely, or append if no match found
        //     // This ensures we update existing entries rather than creating duplicates
        //     const result = this.copyData(childSheet, ACTIVITY_HEADER_ROW_INDEX, true);

        //     if (result.success) {
        //         Logger.log(`[${operationName}] Successfully copied to child at row ${result.targetRowIndex} - ${result.actionTaken}`);
        //     } else {
        //         Logger.log(`[${operationName}] Failed to copy to child: ${result.message}`);
        //     }

        //     return result;
        // } catch (error) {
        //     Logger.log(`[${operationName}] Error in copyDataToChild: ${error.message}`);
        //     throw error;
        // }

        const childSheet = getChildSpreadsheet(this.activity);
        Logger.log(`[${operationName}] Copying data to child sheet "${childSheet.getName()}"`);

        // Use overwrite = true to update existing rows completely, or append if no match found
        // This ensures we update existing entries rather than creating duplicates
        const result = this.copyData(childSheet, ACTIVITY_HEADER_ROW_INDEX, true);

        if (result.success) {
            Logger.log(`[${operationName}] Successfully copied to child at row ${result.targetRowIndex} - ${result.actionTaken}`);
        } else {
            Logger.log(`[${operationName}] Failed to copy to child: ${result.message}`);
        }

        return result;
    }

    copyDataToMaster({ attempt = 1 } = {}) {
        const operationName = 'copyDataToMaster';
        const maxAttempts = 2;

        try {
            const flushStart = new Date();
            try {
                Logger.log(`[${operationName}] Flushing pending edits before master sync (attempt ${attempt})`);
                SpreadsheetApp.flush();
                Logger.log(`[${operationName}] Flush completed in ${new Date() - flushStart}ms`);
            } catch (flushError) {
                Logger.log(`[${operationName}] Flush warning: ${flushError.message}`);
            }

            // Force refresh to get the latest activity values, especially after updates like PROCESSED_DATE
            const activityValueMap = this.activity.getActivityValueMap(true, true);

            // Enhanced validation and debugging
            Logger.log(`[${operationName}] Activity value map retrieved (forced refresh) - REQUEST_TYPE: "${activityValueMap.REQUEST_TYPE}", REQUEST_NUMBER: "${activityValueMap.REQUEST_NUMBER}", PROCESSED_DATE: "${activityValueMap.PROCESSED_DATE}"`);

            if (!activityValueMap.REQUEST_TYPE) {
                const errorMsg = `REQUEST_TYPE is null/undefined for row ${this.activity.rowIndex}`;
                Logger.log(`[${operationName}] Error: ${errorMsg}`);
                return {
                    targetSheet: null,
                    targetRowIndex: null,
                    rowValues: null,
                    success: false,
                    message: errorMsg
                };
            }

            const targetSheet = getSheetName(
                activityValueMap.REQUEST_TYPE
            );

            if (!targetSheet) {
                const errorMsg = `No target sheet found for REQUEST_TYPE: "${activityValueMap.REQUEST_TYPE}"`;
                Logger.log(`[${operationName}] Error: ${errorMsg}`);
                return {
                    targetSheet: null,
                    targetRowIndex: null,
                    rowValues: null,
                    success: false,
                    message: errorMsg
                };
            }

            const masterSheet = getMasterSpreadsheet(targetSheet);

            if (!masterSheet) {
                const errorMsg = `Master sheet not found for target sheet: "${targetSheet}"`;
                Logger.log(`[${operationName}] Error: ${errorMsg}`);
                return {
                    targetSheet: null,
                    targetRowIndex: null,
                    rowValues: null,
                    success: false,
                    message: errorMsg
                };
            }

            Logger.log(`[${operationName}] Copying data to master sheet "${masterSheet.getName()}" for request type "${activityValueMap.REQUEST_TYPE}" (attempt ${attempt})`);

            // Set overwrite = false for copyDataToMaster to merge with existing data instead of replacing
            const result = this.copyData(masterSheet, ACTIVITY_HEADER_ROW_INDEX, false);

            if (result.success) {
                Logger.log(`[${operationName}] Successfully copied to master at row ${result.targetRowIndex} - ${result.actionTaken}`);
                return result;
            }

            Logger.log(`[${operationName}] Failed to copy to master: ${result.message}`);

            if (this._isSpreadsheetTimeout(result.error || result.message) && attempt < maxAttempts) {
                const waitMs = 500 * attempt;
                Logger.log(`[${operationName}] Detected spreadsheet timeout. Retrying in ${waitMs}ms (attempt ${attempt + 1}/${maxAttempts}).`);
                Utilities.sleep(waitMs);
                return this.copyDataToMaster({ attempt: attempt + 1 });
            }

            return result;
        } catch (error) {
            Logger.log(`[${operationName}] Error in copyDataToMaster: ${error.message}`);
            Logger.log(`[${operationName}] Stack trace: ${error.stack}`);

            if (this._isSpreadsheetTimeout(error) && attempt < maxAttempts) {
                const waitMs = 500 * attempt;
                Logger.log(`[${operationName}] Spreadsheet service timeout detected. Waiting ${waitMs}ms before retry ${attempt + 1}/${maxAttempts}.`);
                Utilities.sleep(waitMs);
                return this.copyDataToMaster({ attempt: attempt + 1 });
            }

            return {
                targetSheet: null,
                targetRowIndex: null,
                rowValues: null,
                success: false,
                message: `Exception: ${error.message}`,
                error: error
            };
        }
    }

    _isSpreadsheetTimeout(errorOrMessage) {
        if (!errorOrMessage) return false;

        let message;
        if (typeof errorOrMessage === 'string') {
            message = errorOrMessage;
        } else if (typeof errorOrMessage.message === 'string') {
            message = errorOrMessage.message;
        } else {
            try {
                message = errorOrMessage.toString();
            } catch (ignored) {
                message = '';
            }
        }

        return message.indexOf('Service Spreadsheets timed out while accessing document') !== -1;
    }

}

function syncDataWithMaster(childActivity, masterActivity) {
    const operationName = 'syncDataWithMaster';

    try {
        const childValueMap = childActivity.getActivityValueMap();
        const masterValueMap = masterActivity.getActivityValueMap();
        const childHeaders = getColumnHeaders(childActivity.sheet, true, ACTIVITY_HEADER_ROW_INDEX);

        const rowValues = childHeaders
            .map(header => !childValueMap[header]
                ? masterValueMap[header]
                : childValueMap[header])
            // Remove trailing null values
            .filter((_, index, arr) => index < arr.findLastIndex(value => value) + 1);

        Logger.log(`[${operationName}] Syncing data between child and master activities`);

        // Use the optimized insertRowValues which includes locking
        const resultRowIndex = insertRowValues(childActivity.sheet, rowValues);

        Logger.log(`[${operationName}] Successfully synced data to row ${resultRowIndex}`);

        return {
            success: true,
            rowIndex: resultRowIndex,
            message: "Data synced successfully"
        };

    } catch (error) {
        Logger.log(`[${operationName}] Error syncing data: ${error.message}`);

        return {
            success: false,
            rowIndex: null,
            message: `Sync failed: ${error.message}`,
            error: error
        };
    }
}

function handleChildCustomTabColor(spreadsheet) {
    const sheetsToCheck = [
        { name: "Overdue", range: "A4:F4", text: "No results found" },
        { name: "Duplicate", range: "A4:L4", text: "No Match" },
        { name: "Error", range: "A4:F4", text: "No Results" },
    ];

    sheetsToCheck.forEach(sheetInfo => {
        const sheet = spreadsheet.getSheetByName(sheetInfo.name);
        if (!sheet) return
        const range = sheet.getRange(sheetInfo.range);
        const values = range.getValues()[0];

        const hasDifferentContent = values.some(cell => cell !== sheetInfo.text);

        hasDifferentContent
            ? sheet.setTabColor("#FF0000")
            : sheet.setTabColor(null);
    });
}
// NOTE : Not Used
function handleHasActiveRequest(sheet) {
    const isActive = hasActiveRequest(sheet);
    let color = isActive ? '#FF0000' : null;
    sheet.setTabColor(color);

    isActive
        ? sheet.showSheet()
        : sheet.hideSheet()
}
