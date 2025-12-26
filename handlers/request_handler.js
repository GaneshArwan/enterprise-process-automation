class RequestHandler {
    constructor(request) {
        this.request = request;
        this.activity = request.activity;
        this.attachment = request.attachment;
        this.email = request.email;
        this.masterConfig = new MasterConfig();
        this.activityHandler = request.activityHandler;
        this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }

    injectNewRequest(additionalValues = {}) {
        const payload = this.request.createPayload();

        const eventPayload = {
            ...payload,
            ...additionalValues
        };

        const event = {
            parameter: { path: '/request' },
            postData: {
                contents: JSON.stringify(eventPayload)
            }
        };

        Logger.log("[InjectNewRequest] Full Payload Data:");
        Logger.log(JSON.stringify(eventPayload, null, 2));
        Logger.log("[InjectNewRequest] Event Data:");
        Logger.log(JSON.stringify(event, null, 2));

        const response = handleRequestSubmission(event);

        try {
            const parsedResponse = JSON.parse(response.getContent());
            Logger.log("Injection Response: " + JSON.stringify(parsedResponse));
            return parsedResponse;
        } catch (parseError) {
            Logger.log("Error parsing response: " + parseError);
            return {
                status: "error",
                message: "Invalid response format",
                originalResponse: response.getContent()
            };
        }
    }

    handleSync(ctx) {
        const { prop, constant, levelOrder } = ctx;
        const valueKey = `RESPON_${prop}`;
        const namekey = `NAME_${prop}`;

        if (!this.activity.hasCol(valueKey)) return;

        let isApprover;
        if (levelOrder > 0) {
            isApprover = this.email.isEmailApprover(ctx);

            const isCol = this.activity.hasCol(`RESPON_${prop}`)
            const isApproverStatus = this.attachment.hasApproverStatus(ctx);
            if (!isCol && !isApproverStatus) return { isLastSequence: true }
            //If Current state is Approver, but no approver registered then return false
        }


        const { [namekey]: ACTIVITY_NAME, [valueKey]: ACTIVITY_STATUS } = this.activity.getActivityValueMap();
        if (isNotEmpty(ACTIVITY_NAME) && isNotEmpty(ACTIVITY_STATUS)) return {
            name: ACTIVITY_NAME, status: ACTIVITY_STATUS, isExist: true
        };

        const [status, name] = this.attachment.getValuesByCell([
            AttachmentValues[`${prop}_STATUS_CELL`], AttachmentValues[`${prop}_NAME_CELL`]
        ]);
        if (!name && !status) return { isApprover };

        const isEmptyName = (isNotEmpty(status) && !name);
        const isValidStatus = isNotEmpty(status) ? Object.values(constant.validStatus).includes(status) : true;
        if (isEmptyName || !isValidStatus) {
            Logger.log(`[HandleSync] Handling Invalid Status at row: ${this.request.rowIndex}`)
            this.attachment.clearValuesByCell([
                AttachmentValues[`${ctx.prop}_STATUS_CELL`]
            ]);
            this.email.sendEmailInvalid(ctx);
            return { isApprover };
        }

        return { name, status, isExist: false, isApprover };
    }

    handleNewSubmission() {
        const startTime = new Date();
        Logger.log(`[OnSubmit-Debug] Starting handleNewSubmission for row ${this.request.rowIndex}`);

        const checkStartTime = new Date();
        const { NEW_SUBMISSION_STATUS } = this.activity.getActivityValueMap();
        if (isNotEmpty(NEW_SUBMISSION_STATUS)) {
            Logger.log(`[OnSubmit-Debug] handleNewSubmission early exit - already sent`);
            return;
        }
        const checkTime = new Date() - checkStartTime;
        Logger.log(`[OnSubmit-Debug] Status check completed in ${checkTime}ms`);

        const emailStartTime = new Date();
        let isSent = false;

        try {
            // Add retry mechanism for email sending
            const maxRetries = 3;
            let retryCount = 0;

            while (!isSent && retryCount < maxRetries) {
                try {
                    isSent = this.email.sendEmailNewRequest();
                    if (!isSent && retryCount < maxRetries - 1) {
                        Logger.log(`[OnSubmit-Debug] Email sending failed, retrying (attempt ${retryCount + 1}/${maxRetries})`);
                        Utilities.sleep(500); // Wait before retry
                    }
                } catch (error) {
                    Logger.log(`[OnSubmit-Debug] Email error on attempt ${retryCount + 1}: ${error.message}`);
                    if (retryCount === maxRetries - 1) {
                        // If all retries failed, just log and continue (don't block the flow)
                        Logger.log(`[OnSubmit-Debug] Email sending failed after ${maxRetries} attempts, but continuing with status update`);
                        isSent = true; // Set to true to allow status update, avoiding infinite loop
                    } else {
                        Utilities.sleep(1000); // Longer wait on error
                    }
                }
                retryCount++;
            }
        } catch (error) {
            Logger.log(`[OnSubmit-Debug] Unexpected email error: ${error.message}`);
            // Continue anyway to prevent blocking
            isSent = true;
        }

        const emailTime = new Date() - emailStartTime;
        Logger.log(`[OnSubmit-Debug] Email operation completed in ${emailTime}ms (${isSent ? 'success' : 'failed'})`);

        if (!isSent) {
            Logger.log(`[OnSubmit-Debug] Email sending failed, but continuing...`);
            return;
        }

        const updateStartTime = new Date();
        this.activity.updateNewSubmissionStatus();
        const updateTime = new Date() - updateStartTime;
        Logger.log(`[OnSubmit-Debug] Status update completed in ${updateTime}ms`);

        const totalTime = new Date() - startTime;
        Logger.log(`[OnSubmit-Debug] handleNewSubmission total: ${totalTime}ms (check=${checkTime}ms, email=${emailTime}ms, update=${updateTime}ms)`);
    }


    clearRequesterValues() {
        this.attachment.clearRequesterValues();
        this.activity.updateRequesterValues(null, null);
    }

    handleAttachmentValidation() {
        if (!this.activity.getAttachment()) return true

        const attachmentValidator = new AttachmentValidator(this.attachment);
        const validationResult = attachmentValidator.execute();

        if (![
            ActivitySheetNames.EXTEND_PIR, ActivitySheetNames.MASTER_SITE,
            ActivitySheetNames.CUSTOMER, ActivitySheetNames.VENDOR
        ]
            .includes(this.request.sheet.getName())
        ) {
            return true
        }

        if (Object.keys(validationResult.EMPTY_MANDATORY).length > 0) {
            Logger.log(`[HandleAttachmentValidation] Processing Sent back for Empty Mandatory Fields with detail: \n${JSON.stringify(validationResult.EMPTY_MANDATORY)}`)

            const reason = attachmentValidator.generateSummary(validationResult)

            const emailFunc = () => this.email.sendEmailSendBackBySystem(reason);
            this.handleRequestSendBackBase(
                emailFunc,
                SystemActor.SYSTEM, reason
            )
            return false
        }

        return true
    }

    /**
     * Validates process status change to ensure required conditions are met.
     * Prevents setting status to 'Completed' when TAKEN_DATE is missing.
     * @returns {boolean} True if the status change is valid, false otherwise.
     */
    validateProcessStatusChange() {
        try {
            const { PROCESS_STATUS, TAKEN_DATE } = this.activity.getActivityValueMap();

            // Check if user is trying to set status to 'Completed' without a taken date
            if (PROCESS_STATUS === MDMStatus.COMPLETED && !TAKEN_DATE) {
                Logger.log(`[ValidateProcessStatus] Preventing 'Completed' status change - TAKEN_DATE is missing for row ${this.request.rowIndex}`);

                // Show user-friendly message and revert the change
                const sheet = this.request.sheet;
                if (sheet && sheet.toast) {
                    sheet.toast(
                        'Cannot set status to "Completed" without a Taken Date. Please set the Taken Date first.',
                        'Status Change Not Allowed',
                        8
                    );
                }

                // Revert the cell value back to the previous status
                // We need to find the process status column and clear the current value
                const columnIndices = this.request.getCachedColumnIndices();
                const processStatusColIndex = columnIndices[ColNames.PROCESS_STATUS];

                if (processStatusColIndex && this.request.rowIndex) {
                    const range = sheet.getRange(this.request.rowIndex, processStatusColIndex);
                    range.setValue(''); // Clear the invalid value
                }

                return false;
            }

            Logger.log(`[ValidateProcessStatus] Status change validation passed for row ${this.request.rowIndex} - PROCESS_STATUS: ${PROCESS_STATUS}, TAKEN_DATE: ${TAKEN_DATE ? 'exists' : 'missing'}`);
            return true;

        } catch (error) {
            Logger.log(`[ValidateProcessStatus] Error during validation for row ${this.request.rowIndex}: ${error.message}`);
            // In case of error, allow the change to proceed to avoid blocking the user
            return true;
        }
    }

    handleAskApproval(ctx) {
        const { prop } = ctx;

        const approvalStatusCol = `ASK_${prop}_STATUS`;
        const { [approvalStatusCol]: status } = this.activity.getActivityValueMap();
        //If status already set or no column in activty
        if (status === undefined) return false;
        if (isNotEmpty(status)) return;

        const isSent = this.email.sendEmailAskApproval(ctx);
        if (!isSent) return

        this.activity.updateValue(approvalStatusCol, getDateNow());
        return true
    }

    handleRequestCompleted(attachmentCtx) {
        const isValidationPassed = this.handleAttachmentValidation()
        if (!isValidationPassed) return

        const { name, status } = attachmentCtx;
        this.activity.updateRequesterValues(status, name);

        return true;
    }

    handleRequestRejected(ctx) {
        this.attachment.protectSpreadsheet();
        this.email.sendEmailRejected(ctx);
        return;
    }

    clearSyncValues(options = {}) {
        // options:
        //   how: 'ALL' | 'BY_INDEX' | 'BY_PROP'
        //   index?: number | number[]
        //   prop?: string | string[]
        const { how = 'ALL', index, prop } = options;
        let targets = [];

        switch (how) {
            case 'BY_INDEX': {
                const idxs = Array.isArray(index) ? index : [index];
                targets = idxs
                    .map(i => ATTACHMENT_SYNC_CONTEXTS[i])
                    .filter(ctx => ctx);
                break;
            }
            case 'BY_PROP': {
                const props = Array.isArray(prop) ? prop : [prop];
                targets = ATTACHMENT_SYNC_CONTEXTS
                    .filter(ctx => props.includes(ctx.prop));
                break;
            }
            case 'ALL':
            default:
                targets = ATTACHMENT_SYNC_CONTEXTS;
        }

        // for each matched context, pull out its prop and call your methods
        targets.forEach(ctx => {
            const p = ctx.prop;
            // build an explicit array of cells
            const cellsToClear = [
                AttachmentValues[`${p}_STATUS_CELL`]
            ];
            // coerce to number, guard against missing prop
            if (Number(ctx.levelOrder) > 0) {
                cellsToClear.push(
                    AttachmentValues[`${p}_NAME_CELL`]
                );
            }

            this.attachment.clearValuesByCell(cellsToClear);
        });
    }

    handleRequestApproved() {
        const sheetName = this.request.sheet.getName();

        Logger.log(`[HandleRequestApproved] Processing approval for row ${this.request.rowIndex}, sheet: ${sheetName}`);

        // Handle total task validation
        if (!this.handleTotalTask()) {
            Logger.log(`[HandleRequestApproved] handleTotalTask failed for row ${this.request.rowIndex}`);
            return false;
        }

        // Clear cache to ensure fresh data for baseline calculation after total task update
        if (this.request.clearActivityValueMapCache) {
            this.request.clearActivityValueMapCache();
        }

        const baselineResult = this.handleBaseline();
        if (!baselineResult) {
            Logger.log(`[HandleRequestApproved] handleBaseline failed for row ${this.request.rowIndex}`);
            return false;
        }
        Logger.log(`[HandleRequestApproved] Baseline set successfully: ${JSON.stringify(baselineResult)}`);

        const allocationResult = this.handleAllocation();
        if (!allocationResult) {
            Logger.log(`[HandleRequestApproved] handleAllocation failed for row ${this.request.rowIndex}`);
            return false;
        }
        Logger.log(`[HandleRequestApproved] Allocation set successfully: ${allocationResult}`);

        // Update workload
        if (baselineResult.estimatedTime) {
            const requestAllocator = new RequestAllocator(this.request);
            requestAllocator.updateMdmWorkload(allocationResult, baselineResult.estimatedTime);
            Logger.log(`[HandleRequestApproved] Workload updated for ${allocationResult} (+${baselineResult.estimatedTime}s)`);
        }

        // Finalize Approval
        Logger.log(`[HandleRequestApproved] Finalizing approval for row ${this.request.rowIndex}`);
        this.attachment.protectSpreadsheet();
        this.email.sendEmailApproved();
        this.activityHandler.copyDataToChild();

        Logger.log(`[HandleRequestApproved] Successfully completed approval process for row ${this.request.rowIndex}`);
        return true;
    }

    //Handle Baseline and Estimated 
    handleBaseline() {
        Logger.log(`[HandleBaseline] Starting baseline calculation for row ${this.request.rowIndex}`);

        // Get request type first - this should be stable
        const activityValueMap = this.activity.getActivityValueMap(true, true); // Force refresh

        const isPromoRequest = activityValueMap.PROMO_TYPE ? true : false;
        const requestTypeKey = isPromoRequest
            ? activityValueMap.PROMO_TYPE
            : activityValueMap.REQUEST_TYPE;

        Logger.log(`[HandleBaseline] RequestType: ${requestTypeKey}`);

        if (!requestTypeKey) {
            Logger.log(`[HandleBaseline] No request type found for row ${this.request.rowIndex}`);
            return false;
        }

        // Get TOTAL_TASK directly from the sheet
        let totalTask = getValueByColumn(
            this.request.sheet,
            isPromoRequest ? ColNames.TOTAL_PROMO : ColNames.TOTAL_TASK,
            this.request.rowIndex,
            ACTIVITY_HEADER_ROW_INDEX
        );

        Logger.log(`[HandleBaseline] Direct sheet read - TotalTask: ${totalTask}`);

        // If direct read fails, fall back to activity value map but force a complete refresh
        if (!totalTask) {
            Logger.log(`[HandleBaseline] Direct read failed, trying complete refresh`);

            if (this.request.clearActivityValueMapCache) {
                this.request.clearActivityValueMapCache();
            }

            invalidateSheetDataCache(this.request.sheet);

            const freshActivityValueMap = this.activity.getActivityValueMap(true, true);
            totalTask = freshActivityValueMap.TOTAL_TASK;

            Logger.log(`[HandleBaseline] Refreshed read - TotalTask: ${totalTask}`);
        }

        if (!totalTask) {
            Logger.log(`[HandleBaseline] No total task found for row ${this.request.rowIndex} after all attempts`);
            return false;
        }

        const { baseline, isTaskBaseline } = this.masterConfig.getBaseline({
            requestType: requestTypeKey,
            totalTask: totalTask,
        });

        if (!baseline) {
            Logger.log(`[HandleBaseline] No baseline configuration found for requestType: ${requestTypeKey}, totalTask: ${totalTask} - continuing without baseline`);
            return true; 
        }
        let estimatedTime = 0;
        if (isTaskBaseline) {
            estimatedTime = baseline * totalTask;
        } else {
            estimatedTime = baseline;
        }
        Logger.log(`[HandleBaseline] Calculated baseline: ${baseline}, estimatedTime: ${estimatedTime} (totalTask: ${totalTask})`);

        const ok = setValuesWithIndexes(
            this.request.sheet,
            [ColNames.BASELINE, ColNames.ESTIMATED_TIME],
            this.request.rowIndex,
            [baseline, estimatedTime]
        );

        if (!ok) {
            Logger.log(`[HandleBaseline] Failed to set baseline values for row ${this.request.rowIndex}`);
            return false;
        }

        Logger.log(`[HandleBaseline] Successfully set baseline values for row ${this.request.rowIndex}`);
        return { baseline, estimatedTime };
    }

    handleAllocation() {
        Logger.log(`[HandleAllocation] Starting allocation for row ${this.request.rowIndex}`);

        const requestAllocator = new RequestAllocator(this.request);
        const processedBy = requestAllocator.allocate();

        if (!processedBy) {
            Logger.log(`[HandleAllocation] No processed by value allocated for row ${this.request.rowIndex}`);
            return false;
        }

        Logger.log(`[HandleAllocation] Allocated to: ${processedBy}`);

        const ok = setValueWithIndex(
            this.request.sheet, ColNames.PROCESSED_BY,
            this.request.rowIndex, processedBy
        );

        if (!ok) {
            Logger.log(`[HandleAllocation] Failed to set processed by value for row ${this.request.rowIndex}`);
            return false;
        }

        Logger.log(`[HandleAllocation] Successfully set processed by for row ${this.request.rowIndex}`);
        return processedBy;
    }


    handleTotalTask() {
        const activityValueMap = this.activity.getActivityValueMap();
        if (activityValueMap.TOTAL_TASK) {
            Logger.log(`[HandleTotalTask] Total task already exists: ${activityValueMap.TOTAL_TASK}`);
            return true;
        }

        Logger.log(`[HandleTotalTask] Validating total task for row ${this.request.rowIndex}`);
        const totalTask = this.attachment.getTotalTask();
        if (!totalTask) {
            Logger.log(`[HandleTotalTask] No total task found, handling empty approved case for row ${this.request.rowIndex}`);
            this.clearRequesterValues();
            this.email.sendEmailEmptyTask();
            return false;
        }

        Logger.log(`[HandleTotalTask] Updating total task value: ${totalTask} for row ${this.request.rowIndex}`);
        const updateResult = this.activity.updateTotalTask(totalTask);

        if (!updateResult) {
            Logger.log(`[HandleTotalTask] Failed to update total task for row ${this.request.rowIndex}`);
            return false;
        }

        if (this.request.clearActivityValueMapCache) {
            this.request.clearActivityValueMapCache();
        }

        Logger.log(`[HandleTotalTask] Successfully updated total task for row ${this.request.rowIndex}`);
        return true;
    }

    handleRequestNoApprover() {
        Logger.log(`[HandleRequestNoApprover] Processing request with no approver for row ${this.request.rowIndex}`);

        const attachmentCtx = {
            status: ApproverStatus.APPROVED,
            name: NO_APPROVER,
            isLastSequence: true,
            prop: ATTACHMENT_SYNC_CONTEXTS[ATTACHMENT_SYNC_CONTEXTS.length - 1].prop
        };

        this.attachment.updateApproverValues(attachmentCtx);
        this.activity.updateApproverValues(attachmentCtx);
        this.activity.updateTimestampEntry(attachmentCtx.prop);

        const approvalResult = this.handleRequestApproved();

        if (approvalResult) {
            Logger.log(`[HandleRequestNoApprover] Successfully processed no approver case for row ${this.request.rowIndex}`);
        } else {
            Logger.log(`[HandleRequestNoApprover] Failed to process no approver case for row ${this.request.rowIndex}`);
        }

        return approvalResult;
    }

    handleRequestExpired() {
        const { RESPON_REQUESTER, TIMESTAMP } = this.activity.getActivityValueMap();
        Logger.log(`Processing request expired on row: ${this.request.rowIndex}, Requester Status: ${RESPON_REQUESTER}, Timestamp: ${TIMESTAMP}`);

        this.activity.updateRequesterValues(
            RequesterStatus.EXPIRED, null
        );

        this.attachment.handleAttachmentExpired();
        this.email.sendEmailExpired();

        Logger.log(`Request expired successfully processed for row: ${this.request.rowIndex}`);
    }

    handleProcessedByTrigger(userEmail) {
        const {
            TAKEN_DATE,
            ATTACHMENT,
            ESTIMATED_TIME_FINISHED,
        } = this.activity.getActivityValueMap();

        if (isNotEmpty(TAKEN_DATE) && isNotEmpty(ESTIMATED_TIME_FINISHED)) return;
        let takenDate = TAKEN_DATE;

        if (!TAKEN_DATE) {
            addDriveEditors(
                extractSheetId(ATTACHMENT),
                [userEmail]
            );
            takenDate = this.activity.updateTakenDate();
        }
        this.activity.updateEstimatedTimeFinished(takenDate);

        this.activityHandler.copyDataToMaster();
    }

    handleRequestSendBackBase(
        emailFunc,                
        actor, reason,
        targetSheet = this.request.sheet,
        targetRowIndex = this.request.rowIndex,
    ) {
        const cleared = clearRowValuesBetweenColumns(
            targetSheet,
            targetRowIndex,
            ColNames.NEW_SUBMISSION_STATUS
        );

        if (!cleared) {
            Logger.log(
                `[HandleRequestSendBackBase] Failed to clear sync values for request on ${targetSheet.getName?.() ?? "unknown sheet"}`
            );
        }

        const RequestClass = getRequestClass(targetSheet.getName());
        const request = new RequestClass(targetSheet, targetRowIndex);

        this.clearSyncValues({ how: "ALL" });
        request.attachment.removeProtection();
        request.activity.updateRequesterValues(RequesterStatus.NEED_REVIEW, null);
        request.requestLogger.addMasterLog(
            ActivityLog.SEND_BACK,
            actor, reason
        )

        if (typeof emailFunc !== "function") {
            Logger.log("[HandleRequestSendBackBase] emailFunc is not a valid function");
            return;
        }

        emailFunc(); 
    }


    handleRequestSendBackApprover(ctx = {}) {
        const [name, reason] = this.attachment.getValuesByCell([
            AttachmentValues[`${ctx.prop}_NAME_CELL`],
            AttachmentValues[`${ctx.prop}_NOTES_CELL`],
        ]);

        const emailFunc = () => this.email.sendEmailSendBackByApprover(name, reason);

        return this.handleRequestSendBackBase(
            emailFunc,
            SystemActor.APPROVER, reason
        );
    }


    handleRequestSendBackMDM() {
        const { REQUEST_NUMBER, REQUEST_TYPE, REMARK } = this.activity.getActivityValueMap();

        if (REQUEST_TYPE === RequestTypes.MERCHANDISE_CREATE_NO_IMAGE) return false;

        const sheetName = getSheetName(REQUEST_TYPE);
        const masterSheet = getMasterSpreadsheet(sheetName);
        const masterRowIndex = getRowIndex(masterSheet, REQUEST_NUMBER);

        const emailFunc = () => this.email.sendEmailSendBackByMDM(REMARK);
        this.handleRequestSendBackBase(
            emailFunc,
            SystemActor.MDM, REMARK,
            masterSheet,
            masterRowIndex
        );

        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        spreadsheet.toast(
            'Request has been sent back and removed from the list.',
            'Request Sent Back', 5
        );

        const sheet = this.request.sheet;
        sheet.deleteRow(this.request.rowIndex);
        return;
    }


    _isBU11() {
        const { PROCESS_STATUS } = this.activity.getActivityValueMap();
        const isBU11 = this.activity.getCompanyName() === "BU11";
        const isMerchandise = this.activity.sheet.getName().includes("Merchandise");
        const isCompleted = PROCESS_STATUS === MDMStatus.COMPLETED;
        if (isBU11 && isMerchandise && isCompleted) {
            this.activity.updateFeedbackStatus();
            return;
        }
    }

    handleProcessStatusTrigger(userEmail = null, previousStatus = null, triggerEmail = true) {
        const {
            PROCESSED_BY,
            PROCESS_STATUS,
            PROCESSED_DATE,
            FEEDBACK_STATUS,
            TAKEN_DATE,
            ATTACHMENT,
            ESTIMATED_TIME_FINISHED
        } = this.activity.getActivityValueMap();

        Logger.log(`[handleProcessStatusTrigger] Values: PROCESSED_BY=${PROCESSED_BY}, PROCESS_STATUS=${PROCESS_STATUS}, PROCESSED_DATE=${PROCESSED_DATE}, FEEDBACK_STATUS=${FEEDBACK_STATUS}, TAKEN_DATE=${TAKEN_DATE}, ATTACHMENT=${ATTACHMENT}, ESTIMATED_TIME_FINISHED=${ESTIMATED_TIME_FINISHED}`);

        if (
            isNotEmpty(PROCESSED_DATE) &&
            isNotEmpty(FEEDBACK_STATUS) &&
            !TAKEN_DATE
        ) {
            Logger.log(`[handleProcessStatusTrigger] Early exit: Both PROCESSED_DATE and FEEDBACK_STATUS are set, but TAKEN_DATE is not.`);
            return;
        }

        if (isNotEmpty(PROCESS_STATUS) && !ESTIMATED_TIME_FINISHED && TAKEN_DATE) {
            this.activity.updateEstimatedTimeFinished(TAKEN_DATE);
        }

        if (PROCESS_STATUS === MDMStatus.ON_GOING) {
            Logger.log(`[handleProcessStatusTrigger] PROCESS_STATUS is ON_GOING. Calling handleProcessedByTrigger.`);
            this.handleProcessedByTrigger(userEmail);
            Logger.log(`[handleProcessStatusTrigger] END (ON_GOING branch)`);
            return;
        }

        if (isNotEmpty(PROCESS_STATUS) &&
            (PROCESSED_BY ? true : previousStatus == MDMStatus.ON_GOING)
        ) {
            Logger.log(`[handleProcessStatusTrigger] Entering main processing branch.`);

            if (PROCESS_STATUS === MDMStatus.SEND_BACK && ATTACHMENT !== "NO ATTACHMENT") {
                Logger.log(`[HandleProcessStatusTrigger] Attempting to handle send back for row: ${this.request.rowIndex}`);
                this.handleRequestSendBackMDM();
                return;
            }

            if (!PROCESSED_DATE) {
                Logger.log(`[handleProcessStatusTrigger] PROCESSED_DATE not set. Updating now.`);
                this.activity.updateProcessedDate();
            }

            if (!FEEDBACK_STATUS && triggerEmail) {
                Logger.log(`[handleProcessStatusTrigger] FEEDBACK_STATUS not set and triggerEmail is true.`);
                
                if (this._isBU11()) {
                    Logger.log(`[handleProcessStatusTrigger] Skipping email because company is BU11`);
                    return;
                }

                Logger.log(`[handleProcessStatusTrigger] Sending processed email...`);
                const isSent = this.email.sendEmailProcessed();
                Logger.log(`[handleProcessStatusTrigger] Email sent: ${isSent}`);
                if (isSent) {
                    Logger.log(`[handleProcessStatusTrigger] Updating FEEDBACK_STATUS.`);
                    this.activity.updateFeedbackStatus();
                }
            }

            addDriveEditors(
                extractSheetId(ATTACHMENT),
                [EMAIL_MDM_GROUP]
            )
            this.activityHandler.copyDataToMaster();
        }

        Logger.log(`[handleProcessStatusTrigger] END`);
    }
}

class RequestHandlerPIRExtend extends RequestHandler {
    constructor(request) {
        super(request);
    }

    handleScriptTypeTrigger() {
        this.activity.updateScriptFile("Generating Script...");

        const values = this.attachment.getValuesBySheet(
            "PIR", AttachmentValues.TASK_START_ROW
        );

        if (!values) {
            this.activity.updateScriptFile("No PIR Values found.");
            return;
        }

        const scriptGenerator = new ScriptGenerator(this.activity);
        const scriptDocURL = scriptGenerator.generateScript(values);

        this.activity.updateScriptFile(scriptDocURL
            ? scriptDocURL
            : "Script Type Undefined");

        this.activityHandler.copyDataToMaster();
    }
}

class RequestHandlerHierarchy extends RequestHandler {
    constructor(request) {
        super(request);
    }

    handleScriptTypeTrigger() {
        this.activity.updateScriptFile("Generating Script...");

        let values = this.attachment.getValuesBySheet(
            "AH Reclass",
            AttachmentValues.TASK_START_ROW
        );

        const base = new Date(getDateNow());
        if (!isNaN(base)) {
            const vf =
                base.getDate() <= 10
                    ? (base.setDate(base.getDate() + 1), base)             
                    : new Date(base.getFullYear(), base.getMonth() + 1, 1); 

            const VALID_FROM =
                `${vf.getFullYear()}${String(vf.getMonth() + 1).padStart(2, "0")}${String(vf.getDate()).padStart(2, "0")}`;

            values.forEach(row => { row.VALID_FROM = VALID_FROM; });
        }

        if (!values) {
            this.activity.updateScriptFile("No AH Reclass Values found.");
            return;
        }

        const scriptGenerator = new ScriptGenerator(this.activity);
        const scriptDocURL = scriptGenerator.generateScript(values);

        this.activity.updateScriptFile(scriptDocURL
            ? scriptDocURL
            : "Script Type Undefined");

        this.activityHandler.copyDataToMaster();
    }
}

class RequestHandlerSite extends RequestHandler {
    constructor(request) {
        super(request);
    }

    handleProcessStatusTrigger(userEmail = null, previousStatus, triggerEmail = false) {

        console.log("PREVIOUS STATUS: ", previousStatus);

        const { PROCESS_STATUS } = this.activity.getActivityValueMap();
        if (PROCESS_STATUS === MDMStatus.REJECTED) {
            Logger.log('Process Status Trigger with Sending Email')
            triggerEmail = true;
            return super.handleProcessStatusTrigger(userEmail, previousStatus, triggerEmail);
        };

        Logger.log('Process Status Trigger without Sending Email')
        return super.handleProcessStatusTrigger(userEmail, previousStatus, triggerEmail);
    }

    handleMdmApprovalDate() {
        const {
            PROCESS_STATUS, PROCESSED_DATE, FEEDBACK_STATUS
        } = this.activity.getActivityValueMap();

        if (isNotEmpty(FEEDBACK_STATUS)) return;

        if (
            isNotEmpty(PROCESS_STATUS) &&
            PROCESS_STATUS !== MDMStatus.REJECTED &&
            isNotEmpty(PROCESSED_DATE)
        ) {
            if (!FEEDBACK_STATUS) {
                const isSent = this.email.sendEmailProcessed();
                if (isSent) {
                    this.activity.updateFeedbackStatus();
                }
            }
        }
    }
}

class RequestHandlerFinance extends RequestHandler {
    constructor(request) {
        super(request);
    }

    handleProcessStatusTrigger(userEmail, previousStatus, triggerEmail = true) {
        super.handleProcessStatusTrigger(
            userEmail, previousStatus, triggerEmail
        );

        const { PROCESS_STATUS, REQUEST_TYPE } = this.activity.getActivityValueMap();

        //If Rejected No New Request Created
        if (PROCESS_STATUS === MDMStatus.REJECTED) return;

        if (REQUEST_TYPE === RequestTypes.COST_CENTER_UNBLOCK_TEMPORARY) {
            const result = this.injectNewRequest({
                requestType: RequestTypes.COST_CENTER_UNBLOCK_BLOCK,
                isApprover: false,
                isApproverII: false,
                isApproverIII: false
            });

            if (result.status === "error") return;
            this.attachment.setLinkBlockUrl(result.data.attachmentUrl);
        }
    }

    handleRequestExpired() {
        const { REQUEST_TYPE, EMAIL_ADDRESS } = this.activity.getActivityValueMap();
        if (REQUEST_TYPE !== RequestTypes.COST_CENTER_UNBLOCK_BLOCK) {
            super.handleRequestExpired();
            return;
        }

        //Handle Auto-Complete for Cost Center Unblock Temporary (Block)
        this.attachment.updateRequesterValues(
            RequesterStatus.COMPLETED, EMAIL_ADDRESS
        );
        this.activity.updateRequesterValues(
            RequesterStatus.COMPLETED, EMAIL_ADDRESS
        );
        this.handleRequestNoApprover();
    }
}

class RequestHandlerCustomer extends RequestHandler {
    constructor(request) {
        super(request);
    }

    handleScriptTypeTrigger() {
        const { REQUEST_TYPE } = this.activity.getActivityValueMap();
        if (![RequestTypes.CUSTOMER_CREATE_BADAN_USAHA, RequestTypes.CUSTOMER_CREATE_PERORANGAN].includes(REQUEST_TYPE)) return;

        try {
            this.activity.updateScriptFile("Generating Script...");

            const values = this.attachment.getValuesBySheet(
                "Customer Create", AttachmentValues.TASK_START_ROW
            );

            if (!values) {
                this.activity.updateScriptFile("No Values found.");
                return;
            }

            const scriptGenerator = new CustomerScriptGenerator(this.activity);
            const scriptDriveUrl = scriptGenerator.generateScript(values);

            this.activity.updateScriptFile(scriptDriveUrl
                ? scriptDriveUrl
                : "Script Type Undefined");

            this.activityHandler.copyDataToMaster();

        } catch (error) {
            this.activity.updateScriptFile(`Err: ${error}`);
        }
    }
}