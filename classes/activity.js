/**
 * Represents an activity within a spreadsheet.
 * * @param {Sheet} sheet The Google Sheets sheet where the activity is located.
 * @param {number} [rowIndex=null] The row index of the activity within the sheet.
 */
class Activity {
    constructor(sheet, rowIndex = null,initialData = null) {
        this.sheet = sheet;
        this.rowIndex = rowIndex;
        this.activityHandler = new ActivityHandler(this);
        this._valueMap = initialData ? this._createValueMap(initialData) : null;
        this._filteredCache = null; // Cache for filtered results
        this._cacheTimestamp = this._valueMap ? Date.now() : null;
    }

    _createValueMap(data) {
        const valueMap = {};
        Object.keys(data).forEach(header => {
            const cleanHeader = upperAndSeparate(header);
            valueMap[cleanHeader] = data[header];
        });
        return valueMap;
    }

    getActivityValueMap(removeDupCols = true, forceRefresh = false) {
        const cacheExpiry = 300000; // 5 minute cache
        const now = Date.now();
        
        // Check if we need to refresh the base cache
        if (forceRefresh || !this._valueMap || 
            !this._cacheTimestamp || (now - this._cacheTimestamp) > cacheExpiry) {
            
            const startTime = new Date();
            this._valueMap = getRowValueMapOptimized(
                this.sheet, this.rowIndex,
                true, ACTIVITY_HEADER_ROW_INDEX
            );
            this._cacheTimestamp = now;
            this._filteredCache = null; // Clear filtered cache
            
            const elapsed = new Date() - startTime;
            Logger.log(`[Activity] Value map refreshed in ${elapsed}ms for row ${this.rowIndex}`);
        } else {
            // Keep cache alive while actively used
            this._cacheTimestamp = now;
        }
        
        if (!removeDupCols) {
            return this._valueMap;
        }

        // Use cached filtered result if available
        if (this._filteredCache) {
            return this._filteredCache;
        }

        // Create and cache filtered map
        const filteredMap = {};
        for (const key in this._valueMap) {
            if (!hasKeywords(key, ['DEPARTMENT_', 'ACCESS_SEQUENCE_'])) {
                filteredMap[key] = this._valueMap[key];
            }
        }
        this._filteredCache = filteredMap;
        return filteredMap;
    }

    /**
     * Retrieves the attachment spreadsheet for the activity.
     */
    getAttachment() {
        const { ATTACHMENT } = this.getActivityValueMap();
        const id = extractSheetId(ATTACHMENT);
        if (!id) return;

        return SpreadsheetApp.openById(id);
    }

    getPromoCode() {
        const { PROMO_TYPE } = this.getActivityValueMap();
        if (!PROMO_TYPE) return
        return extractPromoCode(PROMO_TYPE)
    }

    getAccessSequence() {
        const { ACCESS_SEQUENCE } = this.getActivityValueMap();
        return extractAccessSequence(ACCESS_SEQUENCE)
    }

    getCompanyName() {
        const { COMPANY_CODE_NAME } = this.getActivityValueMap();
        return extractCompanyName(COMPANY_CODE_NAME);
    }

    getBaseName() {
        const sheetName = this.sheet.getName();
        return removeSubmitSuffix(sheetName);
    }
}

/**
 * Extends Activity to include validation methods.
 */
class ActivityWithValidation extends Activity {
    constructor(sheet, rowIndex, initialData = null) {
        super(sheet, rowIndex, initialData);
    }

    hasCol(key) {
        const { [key]: colValue } = this.getActivityValueMap();
        return colValue !== undefined;
    }

    hasRequesterValues() {
        const { RESPON_REQUESTER, NAME_REQUESTER } = this.getActivityValueMap();
        return (isNotEmpty(RESPON_REQUESTER) && isNotEmpty(NAME_REQUESTER))
    }

    hasAccessSequenceValue() {
        return !!this.getActivityValueMap()?.ACCESS_SEQUENCE;
    }

    hasDepartmentValue() {
        return !!this.getActivityValueMap()?.DEPARTMENT;
    }

    hasRequestTypeValue() {
        return !!this.getActivityValueMap()?.REQUEST_TYPE;
    }

    hasValidAttachmentValue() {
        const { ATTACHMENT } = this.getActivityValueMap();
        return ATTACHMENT ? isUrl(ATTACHMENT) : false;
    }

    hasTakenDate() {
        return !!this.getActivityValueMap()?.TAKEN_DATE;
    }

    hasRequestNumber() {
        return !!this.getActivityValueMap()?.REQUEST_NUMBER;
    }
}

/**
 * Extends ActivityWithValidation to include update methods.
 */
class ActivityWithUpdate extends ActivityWithValidation {
    constructor(sheet, rowIndex = null, initialData = null) {
        super(sheet, rowIndex, initialData);
    }
    
    _updateValue(colName, value) {
        const colIndex = handleColParams(this.sheet, colName, ACTIVITY_HEADER_ROW_INDEX);

        if (!colIndex) {
            Logger.log(`[Peringatan] Kolom "${colName}" tidak ditemukan di sheet "${this.sheet.getName()}". Pembaruan dilewati.`);
            return false;
        }

        try {
            this.sheet.getRange(this.rowIndex, colIndex).setValue(value);

            if (this._valueMap) {
                const cacheKey = upperAndSeparate(colName);
                this._valueMap[cacheKey] = value;
                this._filteredCache = null;
                this._cacheTimestamp = Date.now();
            }

            return true;

        } catch (e) {
            Logger.log(`[_updateValue] Gagal menulis ke kolom "${colName}" di baris ${this.rowIndex}. Error: ${e.message}`);
            return false;
        }
    }

    _updateValuesBatch(updates) {
        if (!updates || updates.length === 0) return true;

        const updateData = [];
        const cacheUpdates = [];

        for (const [colName, value] of updates) {
            const colIndex = handleColParams(this.sheet, colName, ACTIVITY_HEADER_ROW_INDEX);
            
            if (!colIndex) {
                Logger.log(`[Warning] Column "${colName}" not found in sheet "${this.sheet.getName()}". Skipping.`);
                continue;
            }

            updateData.push({ colIndex, value, colName });
            
            if (this._valueMap) {
                const cacheKey = upperAndSeparate(colName);
                cacheUpdates.push({ cacheKey, value });
            }
        }

        if (updateData.length === 0) return false;

        try {
            const colIndices = updateData.map(u => u.colIndex);
            const values = updateData.map(u => u.value);
            
            setValuesWithIndexes(this.sheet, colIndices, this.rowIndex, values);

            cacheUpdates.forEach(({ cacheKey, value }) => {
                this._valueMap[cacheKey] = value;
            });
            
            this._filteredCache = null;
            this._cacheTimestamp = Date.now();

            Logger.log(`[_updateValuesBatch] Successfully updated ${updateData.length} columns in batch`);
            return true;

        } catch (e) {
            Logger.log(`[_updateValuesBatch] Failed to batch update columns. Error: ${e.message}`);
            return false;
        }
    }

    updateDeptDefault() {
        const requestName = this.getBaseName();
        const defaultValue = DEPARTMENT_DEFAULT_VALUE_MAP[requestName];
        if (!this._updateValue(ColNames.DEPARTMENT, defaultValue)) return;
        return defaultValue;
    }

    updateReqTypeDefault() {
        const requestName = this.getBaseName();
        const defaultValue = REQUEST_TYPE_DEFAULT_VALUE_MAP[requestName];

        this._updateValue(ColNames.REQUEST_TYPE, defaultValue);
    }

    updateTotalTask(totalTask) {
        const { TOTAL_TASK } = this.getActivityValueMap();
        if (isNotEmpty(TOTAL_TASK)) {
            return true;
        }

        const success = this._updateValue(ColNames.TOTAL_TASK, totalTask);
        
        if (success) {
            this._filteredCache = null;
            if (this._valueMap) {
                const cacheKey = upperAndSeparate(ColNames.TOTAL_TASK);
                this._valueMap[cacheKey] = totalTask;
            }
        }
        return success;
    }

    updateRequesterValues(status, name) {
        this._updateValuesBatch([
            [ColNames.RESPON_REQUESTER, status],
            [ColNames.NAME_REQUESTER, name],
            [ColNames.TIMESTAMP_REQUESTER, status ? getDateNow() : null]
        ]);
    }

    updateValue(colName, value) {
        if (!this.hasCol(colName)) return;
        this._updateValue(ColNames[colName], value);
    }

    updateTimestampEntry(prop) {
        const activityValueMap = this.getActivityValueMap(true, true)
        const { [`TIMESTAMP_${prop}`]: timestamp } = activityValueMap;
        this._updateValue(ColNames.TIMESTAMP_ENTRY, timestamp);
    }

    updateApproverValues(ctx) {
        let { prop, status, name } = ctx;
        if (!this.hasCol(`RESPON_${prop}`)) return;

        const statusCol = ColNames[`RESPON_${prop}`];
        const nameCol = ColNames[`NAME_${prop}`];
        const timestampCol = ColNames[`TIMESTAMP_${prop}`];

        this._updateValuesBatch([
            [statusCol, status],
            [nameCol, name],
            [timestampCol, status ? getDateNow() : null]
        ]);
    }

    updateDocumentNumber(value) {
        this._updateValue(ColNames.DOCUMENT_NUMBER, value);
    }

    updateTakenDate(value = getDateNow()) {
        this._updateValue(ColNames.TAKEN_DATE, value);
        return value;
    }

    updateEstimatedTime(estimatedTimeSeconds) {
        return this._updateValue(ColNames.ESTIMATED_TIME, estimatedTimeSeconds);
    }

    updateEstimatedTimeFinished(takenDate) {
        const { ESTIMATED_TIME, ESTIMATED_TIME_FINISHED } = this.getActivityValueMap();
        if (isNotEmpty(ESTIMATED_TIME_FINISHED) || !ESTIMATED_TIME) return;

        let remainingSeconds = Number(ESTIMATED_TIME);
        if (isNaN(remainingSeconds) || remainingSeconds <= 0) {
            return;
        }

        const WORK_START_HOUR = 9;
        const WORK_END_HOUR = 18;
        const LUNCH_START_HOUR = 12;
        const LUNCH_END_HOUR = 13;
        const SECONDS_PER_WORK_DAY = (WORK_END_HOUR - WORK_START_HOUR - (LUNCH_END_HOUR - LUNCH_START_HOUR)) * 3600;

        let currentDate = parseMDYHMS(takenDate);

        const advanceToNextWorkday = (date) => {
            date.setDate(date.getDate() + 1);
            date.setHours(WORK_START_HOUR, 0, 0, 0);
            while (isWeekend(date) || isHoliday(date)) {
                date.setDate(date.getDate() + 1);
            }
            return date;
        };

        if (currentDate.getHours() >= WORK_END_HOUR || isWeekend(currentDate) || isHoliday(currentDate)) {
            currentDate = advanceToNextWorkday(currentDate);
        } else if (currentDate.getHours() < WORK_START_HOUR) {
            currentDate.setHours(WORK_START_HOUR, 0, 0, 0);
        } else if (currentDate.getHours() >= LUNCH_START_HOUR && currentDate.getHours() < LUNCH_END_HOUR) {
            currentDate.setHours(LUNCH_END_HOUR, 0, 0, 0);
        }
        
        const lunchStart = new Date(currentDate).setHours(LUNCH_START_HOUR, 0, 0, 0);
        const lunchEnd = new Date(currentDate).setHours(LUNCH_END_HOUR, 0, 0, 0);
        const workEnd = new Date(currentDate).setHours(WORK_END_HOUR, 0, 0, 0);
        
        let availableSecondsToday = (workEnd - currentDate.getTime()) / 1000;
        if (currentDate.getTime() < lunchEnd) {
            const lunchDuration = (lunchEnd - Math.max(currentDate.getTime(), lunchStart)) / 1000;
            availableSecondsToday -= Math.max(0, lunchDuration);
        }

        if (remainingSeconds <= availableSecondsToday) {
            currentDate = addTime(currentDate, { seconds: remainingSeconds });
            if (currentDate.getTime() > lunchStart && currentDate.getTime() < lunchEnd) {
                const overflow = (currentDate.getTime() - lunchStart) / 1000;
                currentDate.setTime(lunchEnd);
                currentDate = addTime(currentDate, { seconds: overflow });
            }
        } else {
            remainingSeconds -= availableSecondsToday;
            const fullDaysNeeded = Math.floor(remainingSeconds / SECONDS_PER_WORK_DAY);
            for (let i = 0; i < fullDaysNeeded; i++) {
                currentDate = advanceToNextWorkday(currentDate);
            }
            remainingSeconds %= SECONDS_PER_WORK_DAY;
            currentDate = advanceToNextWorkday(currentDate);
            currentDate = addTime(currentDate, { seconds: remainingSeconds });
            if (currentDate.getHours() >= LUNCH_START_HOUR) {
                currentDate = addTime(currentDate, { hours: 1 });
            }
        }

        const out = getDateNow(currentDate);
        this._updateValue(ColNames.ESTIMATED_TIME_FINISHED,out)
        return out;
    }

    updateRemark(value) {
        this._updateValue(ColNames.REMARK, value);
    }

    updateProcessStatus(value) {
        this._updateValue(ColNames.PROCESS_STATUS, value);
    }

    updateProcessedBy(value) {
        this._updateValue(ColNames.PROCESSED_BY, value);
    }

    updateProcessedDate() {
        this._updateValue(ColNames.PROCESSED_DATE, getDateNow());
    }

    updateNewSubmissionStatus(value = getDateNow()) {
        this._updateValue(ColNames.NEW_SUBMISSION_STATUS, value);
    }

    updateAskApprovalStatus(value = getDateNow()) {
        this._updateValue(ColNames.ASK_APPROVAL_STATUS, value);
    }

    updateAskApprovalFinalStatus(value = getDateNow()) {
        this._updateValue(ColNames.ASK_APPROVAL_FINAL_STATUS, value);
    }

    updateMdmApprovalDate() {
        this._updateValue(ColNames.MDM_APPROVAL_DATE, getDateNow());
    }

    updateFeedbackStatus() {
        this._updateValue(ColNames.FEEDBACK_STATUS, "Email Terkirim");
    }

    updateScriptFile(value) {
        this._updateValue(ColNames.SCRIPT_FILE, value);
    }

    updateRequestNumber(value) {
        this._updateValue(ColNames.REQUEST_NUMBER, value);
    }

    updateCacheOnly(colName, value) {
        if (!this._valueMap) return;
        const cacheKey = typeof colName === 'number'
            ? colName
            : upperAndSeparate(colName);
        this._valueMap[cacheKey] = value;
        this._filteredCache = null;
        this._cacheTimestamp = Date.now();
    }

    updateAttachment(value) {
        this._updateValue(ColNames.ATTACHMENT, value);
    }

    updateDepartment(deptValue, targetColIndex = null) {
        this._updateValue(ColNames.DEPARTMENT, deptValue);

        if (!targetColIndex) return;
        this._updateValue(targetColIndex, null); 
    }

    updateAccessSequence(value, targetColIndex) {
        this._updateValue(ColNames.ACCESS_SEQUENCE, value);
        this._updateValue(targetColIndex, null);
    }

    updateEmailApprover(value = null) {
        this._updateValue(ColNames.EMAIL_APPROVER, value || '-');
    }

    updateAdditionalAttachment(value = null) {
        this._updateValue(ColNames.ADDITIONAL_ATTACHMENT, value ? value : '-');
    }
}