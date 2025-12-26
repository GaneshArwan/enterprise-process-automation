class Request {
    constructor(sheet, rowIndex, colIndex = null, rowData = null) {
        this.sheet = sheet;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.activity = new ActivityWithUpdate(this.sheet, this.rowIndex, rowData);
        this.attachment = new AttachmentWithValidation(this.sheet, this.rowIndex);
        this.attachmentValues = null;
        this.templateUrl = null;
        this.email = new EmailHandler(this.sheet, this.rowIndex);
        this.activityHandler = this.activity.activityHandler;
        this.requestHandler = new RequestHandler(this);
        this.requestLogger = new RequestLogger(this);

        // Cache for activity value map to avoid multiple expensive calls
        this._cachedActivityValueMap = null;
        this._activityValueMapCacheKey = null;

        // Cache for column indices to avoid repeated lookups
        this._cachedColumnIndices = null;

        // Cache for request handler sync results to avoid repeated processing
        this._cachedSyncResults = new Map();

        if (rowData) {
            // Hydrate caches with the data that we already fetched in onEdit to avoid an extra sheet read
            const hydratedMap = {};
            Object.keys(rowData).forEach((header) => {
                const normalizedKey = this._normalizeColumnKey(header);
                if (!normalizedKey) return;
                const rawValue = rowData[header];
                hydratedMap[normalizedKey] = (rawValue === '' || rawValue === undefined) ? null : rawValue;
            });

            this._cachedActivityValueMap = hydratedMap;
            this._activityValueMapCacheKey = 'false_false';

            this.activity._valueMap = hydratedMap;
            this.activity._filteredCache = null;
            this.activity._cacheTimestamp = Date.now();
        }
    }

    /**
     * Get cached activity value map to avoid expensive repeated calls
     */
    getCachedActivityValueMap(refresh = false, withRequestNumber = false) {
        const removeDupCols = Boolean(refresh);
        const needsRequestNumber = Boolean(withRequestNumber);
        const cacheKey = `${removeDupCols}_${needsRequestNumber}`;
        const requestNumberKey = this._normalizeColumnKey(ColNames.REQUEST_NUMBER);
        const hasRequestNumberInCache = Boolean(
            this._cachedActivityValueMap &&
            requestNumberKey &&
            this._cachedActivityValueMap[requestNumberKey]
        );
        const forceRefresh = needsRequestNumber && !hasRequestNumberInCache;

        if (!this._cachedActivityValueMap ||
            this._activityValueMapCacheKey !== cacheKey ||
            forceRefresh) {
            Logger.log(`[OnSubmit-Debug] Loading activity value map (refresh=${refresh}, withRequestNumber=${withRequestNumber})`);
            const startTime = new Date();
            this._cachedActivityValueMap = this.activity.getActivityValueMap(removeDupCols, forceRefresh);
            this._activityValueMapCacheKey = cacheKey;
            const loadTime = new Date() - startTime;
            Logger.log(`[OnSubmit-Debug] Activity value map loaded in ${loadTime}ms`);
        }

        return this._cachedActivityValueMap;
    }

    /**
     * Clear the cached activity value map (call this when data changes)
     */
    clearActivityValueMapCache() {
        this._cachedActivityValueMap = null;
        this._activityValueMapCacheKey = null;
        // Also clear sync cache when activity data changes
        this._cachedSyncResults.clear();
    }

    _normalizeColumnKey(colName) {
        if (typeof colName === 'number') return colName;
        if (!colName) return colName;
        // Already normalized keys (REQUEST_NUMBER etc) are returned as-is
        if (colName === colName.toUpperCase() && !colName.includes(' ')) {
            return colName;
        }
        return upperAndSeparate(colName);
    }

    _updateCachedActivityValue(colName, value) {
        if (!this._cachedActivityValueMap) return;
        const key = this._normalizeColumnKey(colName);
        if (!key) return;
        this._cachedActivityValueMap[key] = value;
    }

    /**
     * Get cached column indices to avoid repeated expensive lookups
     */
    getCachedColumnIndices() {
        if (!this._cachedColumnIndices) {
            Logger.log(`[OnEdit-Debug] Loading column indices cache`);
            const startTime = new Date();
            this._cachedColumnIndices = {
                [ColNames.PROCESSED_BY]: getColumnIndex(this.activity.sheet, ColNames.PROCESSED_BY, ACTIVITY_HEADER_ROW_INDEX),
                [ColNames.PROCESS_STATUS]: getColumnIndex(this.activity.sheet, ColNames.PROCESS_STATUS, ACTIVITY_HEADER_ROW_INDEX),
                [ColNames.SCRIPT_TYPE]: getColumnIndex(this.activity.sheet, ColNames.SCRIPT_TYPE, ACTIVITY_HEADER_ROW_INDEX)
            };
            const loadTime = new Date() - startTime;
            Logger.log(`[OnEdit-Debug] Column indices loaded in ${loadTime}ms: ${JSON.stringify(this._cachedColumnIndices)}`);
        }
        return this._cachedColumnIndices;
    }

    /**
     * Optimize sync result caching for interval processing
     */
    getCachedSyncResult(ctx) {
        const cacheKey = `${ctx.prop}_${ctx.levelOrder}`;

        if (this._cachedSyncResults.has(cacheKey)) {
            Logger.log(`[OnInterval-Debug] Using cached sync result for ${ctx.prop}`);
            return this._cachedSyncResults.get(cacheKey);
        }

        const startTime = new Date();
        const result = this.requestHandler.handleSync(ctx);
        const syncTime = new Date() - startTime;
        Logger.log(`[OnInterval-Debug] handleSync for ${ctx.prop} took ${syncTime}ms`);

        // Cache the result for subsequent calls
        this._cachedSyncResults.set(cacheKey, result);

        return result;
    }

    handleTaskRowMigration() {
        const MIGRATION_DATE_MS = Date.parse('2025-05-07T16:50:00+07:00');
        const { TIMESTAMP } = this.getCachedActivityValueMap();
        const timestampMs = Date.parse(TIMESTAMP);

        if (!Number.isNaN(timestampMs) && timestampMs >= MIGRATION_DATE_MS) {
            AttachmentValues = Object.freeze({
                ...AttachmentValues,
                TASK_START_ROW: 25
            });
        }
    }

    createPayload() {
        const valueMap = this.getCachedActivityValueMap();

        return {
            requestType: valueMap.REQUEST_TYPE,
            emailAddress: valueMap.EMAIL_ADDRESS,
            companyCode: extractCompanyCode(valueMap.COMPANY_CODE_NAME),
            companyName: extractCompanyName(valueMap.COMPANY_CODE_NAME),
            department: valueMap.DEPARTMENT,
            attachmentUrl: valueMap.ATTACHMENT,
            documentNumber: valueMap.DOCUMENT_NUMBER,
            additionalAttachment: valueMap.ADDITIONAL_ATTACHMENT,
            validTo: valueMap.VALID_TO,
            validFrom: valueMap.VALID_FROM,
            promoType: valueMap.PROMO_TYPE,
            totalTask: valueMap.TOTAL_TASK,
            modifyType: valueMap.MODIFY_TYPE,
            byPhoneConfirmation: valueMap.BY_PHONE_CONFIRMATION,
            transactionSection: valueMap.TRANSACTION_SECTION,
            updateTo: valueMap.UPDATE_TO,
            bankType: valueMap.BANK_TYPE,
            totalPromo: valueMap.TOTAL_PROMO
        };
    }

    handleDocumentNumber() {
        const { DOCUMENT_NUMBER } = this.getCachedActivityValueMap();
        if (DOCUMENT_NUMBER === undefined) {
            return true;
        }
        const formatted = DOCUMENT_NUMBER
            .split(',')
            .map(n => n.trim())
            .join('\n');
        this.activity.updateDocumentNumber(formatted);
        this._updateCachedActivityValue(ColNames.DOCUMENT_NUMBER, formatted);
        this.clearActivityValueMapCache(); // Clear cache after update
        return true;
    }

    handleAdditionalAttachment() {
        const startTime = new Date();
        Logger.log(`[OnSubmit-Debug] Starting handleAdditionalAttachment for row ${this.rowIndex}`);

        // Quick check if ADDITIONAL_ATTACHMENT column even exists to avoid expensive cache refresh
        const stepStart0 = new Date();
        const hasAdditionalAttachmentCol = this.activity.hasCol('ADDITIONAL_ATTACHMENT');
        const colCheckTime = new Date() - stepStart0;
        Logger.log(`[OnSubmit-Debug] Column existence check completed in ${colCheckTime}ms`);

        if (!hasAdditionalAttachmentCol) {
            Logger.log(`[OnSubmit-Debug] handleAdditionalAttachment early exit - ADDITIONAL_ATTACHMENT column does not exist`);
            return;
        }

        const stepStart1 = new Date();
        const { ADDITIONAL_ATTACHMENT, REQUEST_NUMBER } = this.getCachedActivityValueMap(false, true);
        const getValueMapTime = new Date() - stepStart1;
        Logger.log(`[OnSubmit-Debug] getCachedActivityValueMap completed in ${getValueMapTime}ms`);

        // Check if REQUEST_NUMBER is available
        if (!REQUEST_NUMBER) {
            Logger.log(`[OnSubmit-Debug] handleAdditionalAttachment early exit - REQUEST_NUMBER not available yet`);
            return;
        }

        if (ADDITIONAL_ATTACHMENT === undefined || isFolderUrl(ADDITIONAL_ATTACHMENT)) {
            Logger.log(`[OnSubmit-Debug] handleAdditionalAttachment early exit - no processing needed`);
            return;
        }

        const stepStart2 = new Date();
        const sheetName = this.activity.getBaseName();
        const companyName = this.activity.getCompanyName();
        const getNameTime = new Date() - stepStart2;
        Logger.log(`[OnSubmit-Debug] getName operations completed in ${getNameTime}ms`);

        const stepStart3 = new Date();
        const { additionalDrive } = getDriveFolder(sheetName, companyName);
        const getDriveFolderTime = new Date() - stepStart3;
        Logger.log(`[OnSubmit-Debug] getDriveFolder completed in ${getDriveFolderTime}ms`);

        if (!additionalDrive) return;

        const stepStart4 = new Date();
        // Ensure REQUEST_NUMBER is not null before creating folder
        const folderName = REQUEST_NUMBER || `TEMP_${Date.now()}`;
        const additionalFolder = additionalDrive.createFolder(folderName);
        const createFolderTime = new Date() - stepStart4;
        Logger.log(`[OnSubmit-Debug] createFolder completed in ${createFolderTime}ms using name: ${folderName}`);

        const stepStart5 = new Date();
        addDriveEditors(additionalFolder.getId(), [...this.email.getAllEmails(), EMAIL_MDM_GROUP]);
        const addEditorsTime = new Date() - stepStart5;
        Logger.log(`[OnSubmit-Debug] addDriveEditors completed in ${addEditorsTime}ms`);

        // Update the Additional Attachment to the new folder URL
        const stepStart6 = new Date();
        const updatedValue = additionalFolder.getUrl();
        this.activity.updateAdditionalAttachment(updatedValue);
        this._updateCachedActivityValue(ColNames.ADDITIONAL_ATTACHMENT, updatedValue);
        this.clearActivityValueMapCache(); // Clear cache after update
        const updateAttachmentTime = new Date() - stepStart6;
        Logger.log(`[OnSubmit-Debug] updateAdditionalAttachment completed in ${updateAttachmentTime}ms`);

        if (!ADDITIONAL_ATTACHMENT) return

        const stepStart7 = new Date();
        const links = ADDITIONAL_ATTACHMENT.split(',').map(link => link.trim());
        links.forEach(link => {
            try {
                if (isFolderUrl(link)) {
                    console.log(`Skipping folder URL: ${link}`);
                    return;  // Skip folder links
                }

                const fileId = extractFileIdFromUrl(link);
                const file = DriveApp.getFileById(fileId);
                additionalFolder.addFile(file);

                const oldParents = file.getParents();
                while (oldParents.hasNext()) {
                    const parent = oldParents.next();
                    parent.removeFile(file);
                }

            } catch (e) {
                console.error(`Failed to move file from link ${link}: ${e.message}`);
            }
        });
        const moveFilesTime = new Date() - stepStart7;
        Logger.log(`[OnSubmit-Debug] moveFiles completed in ${moveFilesTime}ms`);

        const totalTime = new Date() - startTime;
        Logger.log(`[OnSubmit-Debug] handleAdditionalAttachment breakdown: getValueMap=${getValueMapTime}ms, getName=${getNameTime}ms, getDriveFolder=${getDriveFolderTime}ms, createFolder=${createFolderTime}ms, addEditors=${addEditorsTime}ms, updateAttachment=${updateAttachmentTime}ms, moveFiles=${moveFilesTime}ms, total=${totalTime}ms`);
    }

    insertValues(values) {
        this.templateUrl = values.ATTACHMENT_URL || null;
        this.attachmentValues = values.ATTACHMENT_VALUES || null;

        const headers = getColumnHeaders(this.sheet, true, ACTIVITY_HEADER_ROW_INDEX);
        const mapData = headers.map(header => values[header] || '');

        insertRowValues(this.sheet, mapData, this.rowIndex);

        Logger.log(`[OnSubmit] Values inserted to row: ${this.rowIndex}`);
    }

    handleEmailApprover() {
        if (!this.getCachedActivityValueMap().EMAIL_APPROVER) {
            this.activity.updateEmailApprover();
            this.clearActivityValueMapCache(); // Clear cache after update
        }
        return true;
    }

    /**
     * Handles the department column by updating or defaulting the department value.
     * Optimized to avoid unnecessary column lookups and batch operations
     * 
     * @returns {string} The department value if updated or defaulted, otherwise undefined.
     */
    handleDepartmentColumn() {
        const startTime = new Date();
        Logger.log(`[OnSubmit-Debug] Starting handleDepartmentColumn for row ${this.rowIndex}`);

        const map = this.getCachedActivityValueMap(false);
        if (map.DEPARTMENT === undefined || this.activity.hasDepartmentValue()) {
            Logger.log(`[OnSubmit-Debug] handleDepartmentColumn early exit: DEPARTMENT=${map.DEPARTMENT}, hasDepartmentValue=${this.activity.hasDepartmentValue()}`);
            return true;
        }

        const lookupStartTime = new Date();
        const { value: deptValue, colIndex: deptColIndex } = getMultipleColumnValue(
            this.sheet,
            map,
            'DEPARTMENT_'
        );
        const lookupTime = new Date() - lookupStartTime;
        Logger.log(`[OnSubmit-Debug] Department lookup completed in ${lookupTime}ms`);

        if (!deptValue) {
            Logger.log(`[OnSubmit-Debug] No department value found, setting default`);
            const defaultStartTime = new Date();
            const result = Boolean(this.activity.updateDeptDefault());
            if (result) this.clearActivityValueMapCache(); // Clear cache after update
            const defaultTime = new Date() - defaultStartTime;
            Logger.log(`[OnSubmit-Debug] Department default set in ${defaultTime}ms`);
            return result;
        }

        const updateStartTime = new Date();
        this.activity.updateDepartment(deptValue, deptColIndex);
        this.clearActivityValueMapCache(); // Clear cache after update
        const updateTime = new Date() - updateStartTime;
        Logger.log(`[OnSubmit-Debug] Department update completed in ${updateTime}ms`);

        const totalTime = new Date() - startTime;
        Logger.log(`[OnSubmit-Debug] handleDepartmentColumn total: ${totalTime}ms (lookup=${lookupTime}ms, update=${updateTime}ms)`);
        return true;
    }

    /**
     * Handles the request type column by defaulting the request type value if not present.
     */
    handleRequestTypeColumn() {
        if (!this.activity.hasRequestTypeValue()) {
            this.activity.updateReqTypeDefault();
            this.clearActivityValueMapCache(); // Clear cache after update
        }
        return true;
    }

    /**
     * Generates and sets a request number for the request.
     * Optimized with batch operations and improved error handling
     */
    handleRequestNumber() {
        if (this.activity.hasRequestNumber()) return true;

        const startTime = new Date();
        Logger.log(`[OnSubmit-Debug] Starting handleRequestNumber for row ${this.rowIndex}`);

        // Cache the base values to avoid repeated calls
        const stepStart1 = new Date();
        const baseName = this.activity.getBaseName();
        const baseNameTime = new Date() - stepStart1;
        Logger.log(`[OnSubmit-Debug] getBaseName completed in ${baseNameTime}ms`);

        const stepStart2 = new Date();
        const companyName = this.activity.getCompanyName();
        const companyNameTime = new Date() - stepStart2;
        Logger.log(`[OnSubmit-Debug] getCompanyName completed in ${companyNameTime}ms`);

        const stepStart3 = new Date();
        // Add detailed logging for request number generation bottleneck
        Logger.log(`[OnSubmit-Debug] Calling generateRequestNumber with baseName: ${baseName}, companyName: ${companyName}`);
        const requestNumber = generateRequestNumber(baseName, companyName);
        const generateTime = new Date() - stepStart3;
        Logger.log(`[OnSubmit-Debug] generateRequestNumber completed in ${generateTime}ms`);

        if (!requestNumber) {
            console.error(`[Request] Failed to generate request number for row ${this.rowIndex}`);
            return false;
        }

        Logger.log(`[OnSubmit] Request Number generated: ${requestNumber}`);

        const stepStart4 = new Date();
        // Optimized sheet write with retry mechanism for large sheets
        Logger.log(`[OnSubmit-Debug] Setting request number ${requestNumber} at row ${this.rowIndex}`);

        let setResult = false;
        let retryCount = 0;
        const maxRetries = 3;

        while (!setResult && retryCount < maxRetries) {
            try {
                setResult = setValueWithIndex(
                    this.sheet, ColNames.REQUEST_NUMBER,
                    this.rowIndex, requestNumber
                );
                if (!setResult && retryCount < maxRetries - 1) {
                    Logger.log(`[OnSubmit-Debug] setValueWithIndex failed, retrying (attempt ${retryCount + 1}/${maxRetries})`);
                    // Brief pause before retry
                    Utilities.sleep(100);
                }
            } catch (error) {
                Logger.log(`[OnSubmit-Debug] setValueWithIndex error on attempt ${retryCount + 1}: ${error.message}`);
                if (retryCount === maxRetries - 1) {
                    throw error;
                }
                Utilities.sleep(200); // Longer pause on error
            }
            retryCount++;
        }

        const setValueTime = new Date() - stepStart4;
        Logger.log(`[OnSubmit-Debug] setValueWithIndex completed in ${setValueTime}ms (${retryCount} attempts)`);

        const totalTime = new Date() - startTime;
        Logger.log(`[OnSubmit-Debug] handleRequestNumber total breakdown: getBaseName=${baseNameTime}ms, getCompanyName=${companyNameTime}ms, generateRequestNumber=${generateTime}ms, setValueWithIndex=${setValueTime}ms, total=${totalTime}ms`);

        if (!setResult) {
            Logger.log(`[OnSubmit-Debug] handleRequestNumber failed after ${maxRetries} attempts`);
            return false;
        }

        // Sync caches so subsequent operations can reuse the new request number without refetching the sheet
        this.activity.updateCacheOnly(ColNames.REQUEST_NUMBER, requestNumber);
        this._updateCachedActivityValue(ColNames.REQUEST_NUMBER, requestNumber);
        this.clearActivityValueMapCache();

        return setResult;
    }

    /**
     * Handles the attachment by making a copy and setting the attachment URL.
     * 
     * @param {boolean} [withImageFolder=false] Indicates if an image folder should be included.
     * @param {Cell} [imageCell=null] The cell containing the image.
     * @returns {Attachment} The attachment object if successful, otherwise undefined.
     */
    handleAttachment(withImageFolder = false, imageCell = null) {
        if (this.activity.hasValidAttachmentValue()) {
            return true;
        }

        const { attachment } = this.attachment.makeAttachmentCopy(
            withImageFolder,
            imageCell,
            this.templateUrl,
            this.attachmentValues
        );

        if (!attachment) {
            console.error(`[Request] No attachment found for row ${this.rowIndex}`);
            return false;
        }

        console.log("This attachment: ", this.attachment.getAttachment().getName());

        if (this.templateUrl) {
            this.requestHandler.clearSyncValues({ how: 'ALL' });
        }

        this.activity.updateAttachment(attachment.getUrl());
        this._updateCachedActivityValue(ColNames.ATTACHMENT, attachment.getUrl());
        this.clearActivityValueMapCache(); // Clear cache after update

        const { ADDITIONAL_ATTACHMENT } = this.getCachedActivityValueMap();
        if (ADDITIONAL_ATTACHMENT) {
            this.attachment.setAdditionalAttachmentValues(ADDITIONAL_ATTACHMENT);
        }
        return attachment;
    }

    /**
     * Handles the submission of the request by performing various operations.
     */
    handleOnSubmit() {
        const sheetName = this.sheet.getName();
        const companyName = this.activity.getCompanyName();

        // --- tiny timer helper ---
        const t0 = Date.now();
        let last = t0;
        const mark = (label) => {
            const now = Date.now();
            Logger.log(
                `[OnSubmit] ${sheetName}/${companyName} row ${this.rowIndex} - ${label}: +${now - last}ms (total ${now - t0}ms)`
            );
            last = now;
        };

        Logger.log(`[OnSubmit] START row ${this.rowIndex} at ${new Date(t0).toISOString()}`);

        // Pre-flight
        this.handleTaskRowMigration();
        // best-effort cache warmup
        try { this.getCachedActivityValueMap(); } catch (_) { }
        try {
            const map = this.getCachedActivityValueMap(false);
            if (!map.DEPARTMENT && !this.activity.hasDepartmentValue()) {
                getMultipleColumnValue(this.sheet, map, 'DEPARTMENT_');
            }
        } catch (_) { }
        mark('Pre-flight done');

        // Fast handlers
        this.handleRequestTypeColumn(); mark('RequestTypeColumn');
        this.handleEmailApprover(); mark('EmailApprover');
        this.handleDocumentNumber(); mark('DocumentNumber');

        // Slow/ordered handlers
        this.handleRequestNumber(); mark('RequestNumber');
        this.handleDepartmentColumn(); mark('DepartmentColumn');
        this.handleAdditionalAttachment(); mark('AdditionalAttachment');
        this.handleAttachment(); mark('Attachment');

        this.requestHandler.handleNewSubmission(); mark('handleNewSubmission');
        // const result = this.activityHandler.copyDataSubmissionToMaster(); mark('copyDataSubmissionToMaster');

        Logger.log(`[OnSubmit] DONE row ${this.rowIndex} - TOTAL ${Date.now() - t0}ms`);
        return true;
    }

    handleOnInterval(requestNumber) {
        const startTime = new Date();
        Logger.log(`[OnInterval-Debug] Starting handleOnInterval for row ${this.rowIndex} with requestNumber: ${requestNumber}`);

        let stepStart = new Date();
        this.handleTaskRowMigration();
        const migrationTime = new Date() - stepStart;
        Logger.log(`[OnInterval-Debug] handleTaskRowMigration completed in ${migrationTime}ms`);

        stepStart = new Date();
        // Pre-load cache for better performance
        const { TIMESTAMP, REQUEST_NUMBER, RESPON_REQUESTER } = this.getCachedActivityValueMap();
        const getCacheTime = new Date() - stepStart;
        Logger.log(`[OnInterval-Debug] getCachedActivityValueMap completed in ${getCacheTime}ms`);

        if (REQUEST_NUMBER !== requestNumber) {
            Logger.log(`[OnInterval] Unmatched Request Number, Skipping Request (Expected: ${requestNumber}, Got: ${REQUEST_NUMBER})`)
            return;
        }

        try {
            stepStart = new Date();
            // Fast path: Check expiration early with optimized conditions
            let shouldCheckExpiration = false;
            if (isDateExpired(TIMESTAMP)) {
                // Don't expire if Respon Requester is "Need Review" - but continue with attachment sync
                if (RESPON_REQUESTER === RequesterStatus.NEED_REVIEW) {
                    Logger.log(`[OnInterval] Skipping expiration due to Need Review status at Row:${this.rowIndex}, continuing with attachment sync`)
                    // Don't return here - continue to attachment sync process
                } else {
                    // Enhanced expiration logic: expire if completed but not approved at all levels
                    const hasRequesterValues = this.activity.hasRequesterValues();
                    shouldCheckExpiration = shouldExpireRequest(this.activity, hasRequesterValues, TIMESTAMP);

                    if (shouldCheckExpiration) {
                        Logger.log(`[OnInterval] Handling Request Expired at Row:${this.rowIndex}`)
                        this.requestHandler.handleRequestExpired();
                        const expirationTime = new Date() - stepStart;
                        Logger.log(`[OnInterval-Debug] Request expiration handled in ${expirationTime}ms`);
                        return;
                    }
                }
            }
            const expirationCheckTime = new Date() - stepStart;
            Logger.log(`[OnInterval-Debug] Expiration check completed in ${expirationCheckTime}ms`);

            stepStart = new Date();
            //Get Attachment Contexts - optimized processing with early filtering
            const attachmentContexts = [];
            let contextProcessingStart = new Date();

            for (const ctx of ATTACHMENT_SYNC_CONTEXTS) {
                const ctxStart = new Date();
                // Use cached sync results for better performance
                const attachmentCtx = this.getCachedSyncResult(ctx);
                const ctxTime = new Date() - ctxStart;

                if (attachmentCtx && typeof attachmentCtx === 'object' && Object.keys(attachmentCtx).length > 0) {
                    attachmentContexts.push({ ...attachmentCtx, ...ctx });
                }
            }
            const attachmentContextTime = new Date() - stepStart;
            Logger.log(`[OnInterval-Debug] Attachment contexts processing completed in ${attachmentContextTime}ms (found ${attachmentContexts.length} contexts)`);

            // Early exit if no contexts to process
            if (attachmentContexts.length === 0) {
                Logger.log(`[OnInterval-Debug] No attachment contexts to process, exiting early`);
                return true;
            }

            // Pre-process context properties to avoid repeated calculations
            stepStart = new Date();
            attachmentContexts.forEach((ctx, i) => {
                ctx.isLastSequence = (i === attachmentContexts.length - 1);
                ctx.levelOrder = i;
            });

            for (let i = 0; i < attachmentContexts.length; i++) {
                const iterationStart = new Date();
                const attachmentCtx = attachmentContexts[i];
                const { isExist, isApprover, status } = attachmentCtx;

                Logger.log(`[OnInterval-Debug] Processing context ${i}(${attachmentCtx.prop}): isExist=${isExist}, isApprover=${isApprover}, status=${status}`);

                // Skip if already exists
                if (isExist) {
                    Logger.log(`[OnInterval-Debug] Context ${i} already exists, skipping`);
                    continue;
                }

                // Early break for requester level without status
                if (i === 0 && !status) {
                    Logger.log(`[OnInterval-Debug] Context ${i} is first level without status, breaking`);
                    break;
                }

                // Process based on status and approver state with early returns
                if (!status && isApprover) {
                    Logger.log(`[OnInterval] Handling Ask Approval at Row:${this.rowIndex}`)
                    this.requestHandler.handleAskApproval(attachmentCtx);
                    const iterationTime = new Date() - iterationStart;
                    Logger.log(`[OnInterval-Debug] Ask approval handled in ${iterationTime}ms`);
                    return;
                }

                if (status === RequesterStatus.COMPLETED) {
                    Logger.log(`[OnInterval] Handling Request Completed at Row:${this.rowIndex}`)
                    const completedStart = new Date();
                    const handleCompleted = this.requestHandler.handleRequestCompleted(attachmentCtx);
                    const completedTime = new Date() - completedStart;
                    Logger.log(`[OnInterval-Debug] handleRequestCompleted took ${completedTime}ms`);

                    if (!handleCompleted) {
                        const iterationTime = new Date() - iterationStart;
                        Logger.log(`[OnInterval-Debug] Request completed handling failed in ${iterationTime}ms`);
                        return;
                    }
                    if (!attachmentCtx.isLastSequence) {
                        const iterationTime = new Date() - iterationStart;
                        Logger.log(`[OnInterval-Debug] Request completed (not last sequence) handled in ${iterationTime}ms`);
                        continue;
                    }
                }

                if (status === ApproverStatus.SEND_BACK) {
                    Logger.log(`[OnInterval] Handling Request Send Back at Row:${this.rowIndex}`)
                    this.requestHandler.handleRequestSendBackApprover(attachmentCtx);
                    const iterationTime = new Date() - iterationStart;
                    Logger.log(`[OnInterval-Debug] Send back handled in ${iterationTime}ms`);
                    return;
                }

                // Handle no approver case
                if (isApprover === false) {
                    Logger.log(`[OnInterval] Handling Request No Approver at Row:${this.rowIndex}`)
                    attachmentCtx.status = ApproverStatus.APPROVED;
                    attachmentCtx.name = NO_APPROVER;

                    const updateStart = new Date();
                    this.attachment.updateApproverValues(attachmentCtx);
                    const updateTime = new Date() - updateStart;
                    Logger.log(`[OnInterval-Debug] updateApproverValues took ${updateTime}ms`);
                }

                const activityUpdateStart = new Date();
                this.activity.updateApproverValues(attachmentCtx);
                const activityUpdateTime = new Date() - activityUpdateStart;
                Logger.log(`[OnInterval-Debug] activity.updateApproverValues took ${activityUpdateTime}ms`);

                // Handle rejection
                if (attachmentCtx.status === ApproverStatus.REJECTED) {
                    Logger.log(`[OnInterval] Handling Request Rejected at Row:${this.rowIndex}`)
                    this.requestHandler.handleRequestRejected(attachmentCtx);
                    const iterationTime = new Date() - iterationStart;
                    Logger.log(`[OnInterval-Debug] Request rejection handled in ${iterationTime}ms`);
                    return;
                }

                // Handle final approval
                if (attachmentCtx.isLastSequence) {
                    Logger.log(`[OnInterval] Handling Request On Last Sequence at Row:${this.rowIndex}`)
                    const finalStart = new Date();

                    try {
                        this.activity.updateTimestampEntry(attachmentCtx.prop);
                        const approvalResult = this.requestHandler.handleRequestApproved();

                        if (!approvalResult) {
                            Logger.log(`[OnInterval] handleRequestApproved failed for row ${this.rowIndex}`);
                            // Don't return here - log the error but continue
                        } else {
                            Logger.log(`[OnInterval] handleRequestApproved completed successfully for row ${this.rowIndex}`);
                        }
                    } catch (approvalError) {
                        Logger.log(`[OnInterval] Error in handleRequestApproved for row ${this.rowIndex}: ${approvalError.message}`);
                        console.error(`[OnInterval] Approval error: ${approvalError.toString()}`);
                        // Continue processing despite approval errors
                    }

                    const finalTime = new Date() - finalStart;
                    Logger.log(`[OnInterval-Debug] Final approval took ${finalTime}ms`);
                }

                const iterationTime = new Date() - iterationStart;
                Logger.log(`[OnInterval-Debug] Context ${i} processing completed in ${iterationTime}ms`);
            }

            const contextProcessingTime = new Date() - stepStart;
            Logger.log(`[OnInterval-Debug] All contexts processing completed in ${contextProcessingTime}ms`);

            const totalTime = new Date() - startTime;
            Logger.log(`[OnInterval-Debug] === TIMING SUMMARY for handleOnInterval row ${this.rowIndex} ===`);
            Logger.log(`[OnInterval-Debug] Migration: ${migrationTime}ms`);
            Logger.log(`[OnInterval-Debug] Cache Load: ${getCacheTime}ms`);
            Logger.log(`[OnInterval-Debug] Expiration Check: ${expirationCheckTime}ms`);
            Logger.log(`[OnInterval-Debug] Attachment Contexts: ${attachmentContextTime}ms`);
            Logger.log(`[OnInterval-Debug] Context Processing: ${contextProcessingTime}ms`);
            Logger.log(`[OnInterval-Debug] TOTAL TIME: ${totalTime}ms (${(totalTime / 1000).toFixed(2)}s)`);
            Logger.log(`[OnInterval-Debug] === END TIMING SUMMARY ===`);

            return true;

        } catch (error) {
            const totalTime = new Date() - startTime;
            Logger.log(`[OnInterval-Debug] ERROR after ${totalTime}ms: ${error.message}`);
            console.error(`[handleOninterval] Error processing interval check for row ${this.rowIndex}: ${error.toString()}`);
            throw error;
        }
    }

    /**
     * Handles the request on edit by checking and executing specific column handlers.
     */
    handleOnEdit(userEmail, previousStatus) {
        SpreadsheetApp.flush();
        console.log("PREVIOUS STATUS: ", previousStatus);
        const startTime = new Date();
        Logger.log(`[OnEdit-Debug] Starting handleOnEdit for row ${this.rowIndex}, column ${this.colIndex}, user: ${userEmail}`);

        console.log("RUNNING handleOnEdit");

        let stepStart = new Date();
        this.handleTaskRowMigration();
        const migrationTime = new Date() - stepStart;
        Logger.log(`[OnEdit-Debug] handleTaskRowMigration completed in ${migrationTime}ms`);

        stepStart = new Date();
        const hasRequesterTime = new Date() - stepStart;

        if (this.rowIndex <= ACTIVITY_HEADER_ROW_INDEX) {
            const totalTime = new Date() - startTime;
            console.log(`[handleOnEdit] Skipping row ${this.rowIndex} due to insufficient data (rowIndex <= ${ACTIVITY_HEADER_ROW_INDEX} or no requester/approver)`);
            Logger.log(`[OnEdit-Debug] Early exit after ${totalTime}ms`);
            return;
        }

        stepStart = new Date();
        // Map of column names to their handlers
        const columnHandlers = {
            [ColNames.PROCESSED_BY]: () =>
                this.requestHandler.handleProcessedByTrigger(userEmail),
            [ColNames.PROCESS_STATUS]: () => {
                // Validate process status change before executing handler
                if (this.requestHandler.validateProcessStatusChange()) {
                    this.requestHandler.handleProcessStatusTrigger(userEmail, previousStatus);
                }
            },
            [ColNames.SCRIPT_TYPE]: () =>
                this.requestHandler.handleScriptTypeTrigger()
        };

        // Get cached column indices to avoid repeated expensive lookups
        const columnIndices = this.getCachedColumnIndices();

        const columnIndexTime = new Date() - stepStart;
        Logger.log(`[OnEdit-Debug] Column index lookup completed in ${columnIndexTime}ms`);
        Logger.log(`[OnEdit-Debug] Column indices: ${JSON.stringify(columnIndices)}, target colIndex: ${this.colIndex}`);

        stepStart = new Date();
        // Find and execute the matching handler
        const matchingColumn = Object.entries(columnIndices)
            .find(([_, index]) => index === this.colIndex);

        if (matchingColumn) {
            const handlerName = matchingColumn[0];
            Logger.log(`[OnEdit-Debug] Executing handler for column: ${handlerName}`);
            columnHandlers[handlerName]();
            const handlerTime = new Date() - stepStart;
            Logger.log(`[OnEdit-Debug] Handler "${handlerName}" completed in ${handlerTime}ms`);
        } else {
            Logger.log(`[OnEdit-Debug] No matching handler found for column index ${this.colIndex}`);
        }

        const totalTime = new Date() - startTime;
        Logger.log(`[OnEdit-Debug] === TIMING SUMMARY for handleOnEdit row ${this.rowIndex} ===`);
        Logger.log(`[OnEdit-Debug] Migration: ${migrationTime}ms`);
        Logger.log(`[OnEdit-Debug] Requester Check: ${hasRequesterTime}ms`);
        Logger.log(`[OnEdit-Debug] Column Index Lookup: ${columnIndexTime}ms`);
        Logger.log(`[OnEdit-Debug] Handler Execution: ${(new Date() - stepStart)}ms`);
        Logger.log(`[OnEdit-Debug] TOTAL TIME: ${totalTime}ms (${(totalTime / 1000).toFixed(2)}s)`);
        Logger.log(`[OnEdit-Debug] === END TIMING SUMMARY ===`);
        SpreadsheetApp.flush();
    }

    handleOnChildInterval() {
        const startTime = new Date();
        Logger.log(`[ChildInterval] Starting for row ${this.rowIndex}`);

        const flushWithLog = (reason) => {
            try {
                const flushStart = new Date();
                Logger.log(`[ChildInterval] Flushing pending edits (${reason}) for row ${this.rowIndex}`);
                SpreadsheetApp.flush();
                Logger.log(`[ChildInterval] Flush completed in ${new Date() - flushStart}ms for row ${this.rowIndex}`);
            } catch (flushError) {
                Logger.log(`[ChildInterval] Flush warning (${reason}) for row ${this.rowIndex}: ${flushError.message}`);
            }
        };

        try {
            this.handleTaskRowMigration();

            const {
                PROCESS_STATUS,
                TAKEN_DATE,
                PROCESSED_DATE,
                FEEDBACK_STATUS,
                ESTIMATED_TIME,
                ESTIMATED_TIME_FINISHED,
                MDM_APPROVAL_DATE,
                ATTACHMENT,
                NO_AR_TO_SAP,
            } = this.getCachedActivityValueMap(true);

        // --- PERBAIKAN TERTARGET BERDASARKAN LOGIKA BARU ---

        // Fix 1: Missing Estimated Time Finished
        if (TAKEN_DATE && ESTIMATED_TIME && !ESTIMATED_TIME_FINISHED) {
            Logger.log(`[ChildInterval] Fixing missing Estimated Time Finished for row ${this.rowIndex}`);
            this.activity.updateEstimatedTimeFinished(TAKEN_DATE);
            flushWithLog('estimated time finished recovery');
            return;
        }

        // Fix 2: Failed "Send Back" handling
        if (PROCESS_STATUS === MDMStatus.SEND_BACK && !FEEDBACK_STATUS && ATTACHMENT !== "NO ATTACHMENT") {
            Logger.log(`[ChildInterval] Retrying failed 'Send Back' for row ${this.rowIndex}`);
            this.requestHandler.handleRequestSendBackMDM();
            flushWithLog('retry send back');
            return;
        }

        // Fix 3 & 4 digabung: Penanganan Feedback Status yang hilang
        if (!FEEDBACK_STATUS) {
            // Kondisi untuk alur kerja Master Site
            if (NO_AR_TO_SAP && MDM_APPROVAL_DATE) {
                Logger.log(`[ChildInterval] Fixing missing Feedback Status (Master Site flow) for row ${this.rowIndex}`);
                this.requestHandler.handleMdmApprovalDate();
                flushWithLog('master site feedback recovery');
                return;
            }
            // Kondisi untuk alur kerja standar
            else if (!NO_AR_TO_SAP && PROCESS_STATUS && PROCESS_STATUS !== MDMStatus.ON_GOING && PROCESSED_DATE) {
                Logger.log(`[ChildInterval] Fixing missing Feedback Status (Standard flow) for row ${this.rowIndex}`);
                this.requestHandler.handleProcessStatusTrigger(null, MDMStatus.ON_GOING, true);
                flushWithLog('standard feedback recovery');
                return;
            }
        }

        Logger.log(`[ChildInterval] No specific fix applied for row ${this.rowIndex}, assuming resolved.`);

        } catch (error) {
            const totalTime = new Date() - startTime;
            Logger.log(`[ChildInterval] ERROR in row ${this.rowIndex} after ${totalTime}ms: ${error.message}`);
            console.error(`[handleOnChildInterval] Error processing interval check for row ${this.rowIndex}: ${error.toString()}`);
            throw error;
        }
    }
}

/**
 * Extends the Request class to handle PIR (Purchase Invoice Request) specific operations.
 */
class RequestPIRExtend extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.requestHandler = new RequestHandlerPIRExtend(this);
    }
}

class RequestHierarchy extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.requestHandler = new RequestHandlerHierarchy(this);
    }
}

/**
 * Extends the Request class to handle requests with image attachments.
 */
class RequestWithImage extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
    }

    /**
     * Handles the attachment with image folder.
     * 
     * @returns {Attachment} The attachment object if successful, otherwise undefined.
     */
    handleAttachment() {
        return super.handleAttachment(true, AttachmentValues.IMAGE_CELL);
    }
}

/**
 * Extends the Request class to handle promotion creation requests.
 */
class RequestPromotion extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.attachment = new AttachmentPromotion(sheet, rowIndex); // The attachment object specific to promotion creation.
        this.requestHandler = new RequestHandler(this);
    }

    /**
     * Handles the promotion creation process by creating or deleting sheets based on the total task count.
     * 
     * @param {Attachment} attachment The attachment object associated with the promotion creation.
     */
    _handlePromotionCreate(attachment) {
        const targetSheet = attachment.getSheets()[1];

        const { TOTAL_PROMO } = this.getCachedActivityValueMap();
        if (+TOTAL_PROMO == 1) {
            attachment.deleteSheet(targetSheet);
            return;
        }

        for (let i = 3; i <= TOTAL_PROMO; i++) {
            let newSheet = targetSheet.copyTo(attachment);
            newSheet.setName("PROMO " + i);
            attachment.setActiveSheet(newSheet);
            attachment.moveActiveSheet(attachment.getSheets().length - 1);
        }
    }

    /**rt nb
     * Handles the attachment for promotion creation requests.
     * 
     * @returns {Attachment} The attachment object if successful, otherwise undefined.
     */
    handleAttachment() {
        console.log("[HandleAttachment] Handling attachment for RequestPromotion");
        const attachment = super.handleAttachment();
        const { REQUEST_TYPE } = this.getCachedActivityValueMap();


        if (REQUEST_TYPE == RequestTypes.PROMOTION_CREATE && !this.templateUrl) {
            console.log("Handle Promotion Create")
            this._handlePromotionCreate(attachment);
        }

        return attachment;
    }
}

class RequestMerchandise extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
    }

    handleOnSubmit() {
        //Update requester status to completed by default
        const { targetRowIndex, targetSheet } = super.handleOnSubmit();

        const { REQUEST_TYPE, EMAIL_ADDRESS, REQUEST_NUMBER } = this.getCachedActivityValueMap();
        if (REQUEST_TYPE === RequestTypes.MERCHANDISE_CREATE_NO_IMAGE) {

            this.activity.updateRequesterValues(
                RequesterStatus.COMPLETED,
                EMAIL_ADDRESS
            )

            this.activity.updateTimestampEntry(ATTACHMENT_SYNC_CONTEXTS[0].prop);
            this.requestHandler.handleRequestApproved();
        }
    }

    handleAttachment() {
        const { REQUEST_TYPE } = this.getCachedActivityValueMap();
        if (REQUEST_TYPE === RequestTypes.MERCHANDISE_CREATE_IMAGE) {
            return super.handleAttachment(true, AttachmentValues.IMAGE_CELL);
        }

        Logger.log("[HandleAttachment] Setting attachment to NO ATTACHMENT");
        this.activity.updateAttachment("NO ATTACHMENT");
        this._updateCachedActivityValue(ColNames.ATTACHMENT, "NO ATTACHMENT");
        this.clearActivityValueMapCache(); // Clear cache after update
        return true;
    }
}

class RequestMasterSite extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.requestHandler = new RequestHandlerSite(this);
    }

    handleOnEdit(userEmail, previousStatus) {
        const startTime = new Date();
        Logger.log(`[RequestMasterSite-Debug] Starting handleOnEdit for row ${this.rowIndex}`);

        let stepStart = new Date();
        super.handleOnEdit(userEmail, previousStatus);
        const superTime = new Date() - stepStart;
        Logger.log(`[RequestMasterSite-Debug] super.handleOnEdit completed in ${superTime}ms`);

        stepStart = new Date();
        const mdmApprovalDateIndex = getColumnIndex(
            this.activity.sheet,
            ColNames.MDM_APPROVAL_DATE,
            ACTIVITY_HEADER_ROW_INDEX
        );
        const getIndexTime = new Date() - stepStart;
        Logger.log(`[RequestMasterSite-Debug] MDM approval date index lookup completed in ${getIndexTime}ms`);

        if (this.colIndex === mdmApprovalDateIndex) {
            stepStart = new Date();
            Logger.log("[RequestMasterSite] MDM Approval Date column edited");
            this.requestHandler.handleMdmApprovalDate();
            const handlerTime = new Date() - stepStart;
            Logger.log(`[RequestMasterSite-Debug] handleMdmApprovalDate completed in ${handlerTime}ms`);
        }

        const totalTime = new Date() - startTime;
        Logger.log(`[RequestMasterSite-Debug] Total handleOnEdit time: ${totalTime}ms`);
    }

    handleOnChildInterval() {
        super.handleOnChildInterval();
        this.requestHandler.handleMdmApprovalDate();
    }
}

class RequestMasterFinance extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.attachment = new AttachmentFinanceMaster(sheet, rowIndex);
        this.requestHandler = new RequestHandlerFinance(this);
    }
}

class RequestPricing extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.attachment = new AttachmentPricing(sheet, rowIndex);
    }

    handleAccessSequenceColumn() {
        if (this.getCachedActivityValueMap().ACCESS_SEQUENCE === undefined ||
            this.activity.hasAccessSequenceValue()) {
            return;
        }

        const { value: accSeqVal, colIndex: accSeqColIndex } = getMultipleColumnValue(
            this.sheet,
            this.getCachedActivityValueMap(false),
            'ACCESS_SEQUENCE_'
        );

        this.activity.updateAccessSequence(accSeqVal, accSeqColIndex);
        this._updateCachedActivityValue(ColNames.ACCESS_SEQUENCE, accSeqVal);
        this.clearActivityValueMapCache(); // Clear cache after update
        return accSeqVal;
    }

    handleOnSubmit() {
        this.handleAccessSequenceColumn();
        super.handleOnSubmit();
    }
}

class RequestCustomer extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData);
        this.requestHandler = new RequestHandlerCustomer(this);
    }
}

class RequestVendor extends Request {
    constructor(sheet, rowIndex, colIndex, rowData = null) {
        super(sheet, rowIndex, colIndex, rowData)
        this.email = new EmailHandlerVendor(sheet, rowIndex);
        this.attachment = new AttachmentVendor(sheet, rowIndex);
        this.requestHandler = new RequestHandler(this);
    }
}
