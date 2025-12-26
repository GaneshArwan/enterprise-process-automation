/**
 * Handles POST requests to the web app
 * @param {Object} e - The event object from Apps Script
 * @returns {TextOutput} JSON response
 */
function doPost(e) {
    return handleRequest(e, 'POST');
}

/**
 * Creates a standardized success response
 * @param {Object} data - The data to return
 * @returns {TextOutput} JSON formatted success response
 */
function createSuccessResponse(data) {
    Logger.log("SUCCESS DATA: " + JSON.stringify(data));
    return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: data
    })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Creates a standardized error response
 * @param {string} error - Error message
 * @param {number} [statusCode=400] - HTTP status code
 * @returns {TextOutput} JSON formatted error response
 */
function createErrorResponse(error, statusCode = 400) {
    Logger.log("ERROR DATA: " + error);
    return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: error,
        code: statusCode
    })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Routes requests to appropriate handlers based on method and path
 * @param {Object} e - The event object
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @returns {TextOutput} Response from the handler
 */
function handleRequest(e, method) {
    try {
        Logger.log("[API] Incoming Request...");
        
        // 1. Coba ambil Path dari URL Parameter (?path=/...)
        let path =  e.parameter?.path || '';
        Logger.log(`[API] Path from URL param: "${path}"`);

        // 2. Jika Path kosong, coba intip JSON Body (Payload)
        // Ini untuk mendukung request dari Child Script yang mengirim { action: 'update_workload' }
        if (!path && method === 'POST' && e.postData && e.postData.contents) {
            try {
                Logger.log(`[API] Path empty. Checking JSON Body...`);
                const body = JSON.parse(e.postData.contents);
                Logger.log(`[API] Body Action: "${body.action}"`);

                if (body.action === 'update_workload') {
                    path = '/update_workload'; // Mapping Action ke Path
                    Logger.log(`[API] Path derived from Action: "${path}"`);
                }
            } catch (err) {
                Logger.log(`[API] Error parsing JSON body for path detection: ${err.message}`);
            }
        }

        // 3. Define available endpoints
        const endpoints = {
            POST: {
                '/request': handleRequestSubmission,
                '/update_workload': handleWorkloadUpdate, // Pastikan handler ini ada
            }
        };

        // 4. Validasi Akhir
        if (!endpoints[method] || !endpoints[method][path]) {
            return createErrorResponse(`Invalid endpoint: ${method} ${path}`, 404);
        }

        // 5. Eksekusi Handler
        Logger.log(`[API] Routing to handler for: ${path}`);
        return endpoints[method][path](e);

    } catch (error) {
        Logger.log(`[API Critical Error] ${error.toString()}`);
        return createErrorResponse(error.toString(), 500);
    }
}

/**
 * Handles workload update requests from Child Script
 * @param {Object} e - The event object
 * @returns {TextOutput} JSON response
 */
function handleWorkloadUpdate(e) {
    Logger.log("[handleWorkloadUpdate] Processing update...");
    
    try {
        const payload = JSON.parse(e.postData.contents);

        // Validasi Payload
        if (!payload.mdmName || payload.seconds === undefined) {
            return createErrorResponse("Invalid parameters: mdmName and seconds are required.", 400);
        }

        const mdmName = payload.mdmName;
        const seconds = Number(payload.seconds);

        if (isNaN(seconds)) {
             return createErrorResponse("Invalid parameters: seconds must be a number.", 400);
        }

        // Panggil fungsi internal update property (Pastikan fungsi ini ada di Master/WorkloadManager)
        //
        const newTotal = updateMdmWorkloadProperty(mdmName, seconds);

        Logger.log(`[handleWorkloadUpdate] Success. ${mdmName} new total: ${newTotal}`);

        return createSuccessResponse({
            message: "Workload updated successfully",
            mdmName: mdmName,
            newTotal: newTotal
        });

    } catch (error) {
        Logger.log(`[handleWorkloadUpdate] Error: ${error.message}`);
        return createErrorResponse(error.message, 500);
    }
}


/**
 * Validates request payload for required fields
 * @param {Object} payload - Request payload
 * @returns {Object} Validation result with success flag and error message
 */
function validateRequestPayload(payload) {
    const attachmentSyncKeys = getAttachmentSyncContextKeys();
    attachmentSyncKeys.forEach(key => {
        if (payload[key] === undefined) {
            payload[key] = true;
        }
    });

    const mandatoryFields = {
        requestType: ColNames.REQUEST_TYPE,
        emailAddress: ColNames.EMAIL_ADDRESS,
        companyCode: 'Company Code',
        companyName: 'Company Name',
    };

    // Handle Missing Mandatory Fields
    const missingFields = [];
    for (const [field, label] of Object.entries(mandatoryFields)) {
        if (!payload[field]) {
            missingFields.push(label);
        }
    }

    if (missingFields.length > 0) {
        return {
            isValid: false,
            error: `Missing required fields: ${missingFields.join(', ')}`
        };
    }

    // Validate attachment URL if provided
    if (payload.attachmentUrl && !isUrl(payload.attachmentUrl)) {
        return {
            isValid: false,
            error: 'Attachment is not a valid URL'
        };
    }

    return { isValid: true, payload };
}

/**
 * Prepares values dictionary from payload
 * @param {Object} payload - Validated request payload
 * @returns {Object} Values dictionary for sheet
 */
function prepareValuesDictionary(payload) {
    return {
        TIMESTAMP: getDateNow(),
        REQUEST_TYPE: payload.requestType,
        EMAIL_ADDRESS: payload.emailAddress,
        COMPANY_CODE_NAME: `${payload.companyCode} - ${payload.companyName}`,
        DEPARTMENT: payload.department,
        ...(payload.attachmentUrl && { ATTACHMENT_URL: payload.attachmentUrl }),
        ...(payload.attachmentValues && { ATTACHMENT_VALUES: payload.attachmentValues }),
        ...(payload.documentNumber && { DOCUMENT_NUMBER: payload.documentNumber }),
        ...(payload.additionalAttachment && { ADDITIONAL_ATTACHMENT: payload.additionalAttachment }),
        ...(payload.attachmentUrl && { ATTACHMENT_URL: payload.attachmentUrl }),
        ...(payload.validFrom && { VALID_FROM: payload.validFrom }),
        ...(payload.validTo && { VALID_TO: payload.validTo }),
        ...(payload.promoType && { PROMO_TYPE: payload.promoType }),
        ...(payload.totalTask && { TOTAL_TASK: payload.totalTask }),
        ...(payload.modfiyType && { MODIFY_TYPE: payload.modfiyType }),
        ...(payload.byPhoneConfirmation && { BY_PHONE_CONFIRMATION: payload.byPhoneConfirmation }),
        ...(payload.transactionSection && { TRANSACTION_SECTION: payload.transactionSection }),
        ...(payload.updateTo && { UPDATE_TO: payload.updateTo }),
        ...(payload.bankType && { BANK_TYPE: payload.bankType }),
        ...(payload.totalPromo && { TOTAL_PROMO: payload.totalPromo }),
    };
}

/**
 * Processes request approval
 * @param {Request} request - Request object
 * @param {Object} payload - Request payload
 */
function processRequestSync(request, payload) {
    console.log("PAYLOAD: " + JSON.stringify(payload));

    // Handle requester status based on isRequester flag
    if (payload.isRequester !== false) {
        // Normal submission - mark requester as completed
        request.activity.updateRequesterValues(
            RequesterStatus.COMPLETED,
            payload.requesterName
        );
        request.attachment.updateRequesterValues(
            RequesterStatus.COMPLETED,
            payload.requesterName
        );
    } else {
        console.log("Existing Process Sync: isRequester is false");
        return;
    }

    // Get existing attachment contexts similar to how handleOnInterval does it
    const attachmentContexts = ATTACHMENT_SYNC_CONTEXTS.slice(1).reduce((out, ctx, idx) => {
        const attachmentCtx = request.requestHandler.handleSync(ctx);

        // handleSync returns undefined if the column doesn't exist (!hasCol)
        if (
            attachmentCtx &&
            typeof attachmentCtx === 'object' &&
            Object.keys(attachmentCtx).length > 0
        ) {
            // For send-back scenarios (isRequester=false), determine isApprover from email configuration
            // For normal scenarios, use the value from handleSync or default to false
            let isApprover;
            if (payload.isRequester === false) {
                // Send-back scenario: check if there are actual approvers configured
                isApprover = request.email.isEmailApprover(ctx);
            } else {
                // Normal scenario: use existing logic
                isApprover = attachmentCtx.isApprover !== undefined ? attachmentCtx.isApprover : false;
            }

            out.push({ ...attachmentCtx, ...ctx, originalIndex: idx, isApprover });
        }

        return out;
    }, []);

    const contextKeys = getAttachmentSyncContextKeys().slice(1);
    let hasApprovedStatus = false;
    let lastApprovedContext = null;

    // Process only the contexts that actually exist
    attachmentContexts.forEach((attachmentCtx) => {
        const { originalIndex, prop, isApprover } = attachmentCtx;
        const key = contextKeys[originalIndex];

        if (payload[key] !== undefined) {
            // Payload has explicit approval data for this context
            console.log(`HANDLING APPROVAL FOR: ${key} (prop: ${prop})`);

            // normalize true/false into your enum
            const raw = payload[key];
            const status =
                raw === true
                    ? ApproverStatus.APPROVED
                    : raw === false
                        ? ApproverStatus.REJECTED
                        : raw;

            const name = payload[`${key}Name`] || NO_APPROVER;

            const ctx = { ...attachmentCtx, status, name };
            console.log(`CONTEXT: ${JSON.stringify(ctx)}`);

            // always update both attachment & activity
            request.attachment.updateApproverValues(ctx);
            request.activity.updateApproverValues(ctx);

            if (status === ApproverStatus.APPROVED) {
                hasApprovedStatus = true;
                lastApprovedContext = ctx;
            } else {
                // FIX: correct API usage (clear only this prop)
                request.requestHandler.clearSyncValues({ how: 'BY_PROP', prop: ctx.prop });
            }
        } else if (isApprover === false) {
            // No approver assigned and no payload data
            if (payload.isRequester === false) {
                // SEND-BACK: do not auto-approve; leave cells cleared
                console.log(`SEND-BACK: leaving ${prop} empty (no approver)`);
                // nothing to write; keep hasApprovedStatus as-is
            } else {
                // Normal flow: auto-approve as NO APPROVER
                console.log(`AUTO-APPROVING NO APPROVER FOR: ${prop}`);
                const ctx = {
                    ...attachmentCtx,
                    status: ApproverStatus.APPROVED,
                    name: NO_APPROVER
                };

                request.attachment.updateApproverValues(ctx);
                request.activity.updateApproverValues(ctx);

                hasApprovedStatus = true;
                lastApprovedContext = ctx;
            }
        } else {
            // There are approvers configured but no payload data - skip processing for now
            console.log(`SKIPPING CONTEXT WITH APPROVERS: ${prop} (isApprover: ${isApprover}, hasPayload: false)`);
        }
    });

    // Only call handleRequestApproved once at the end if there were any approvals
    if (hasApprovedStatus) {
        if (lastApprovedContext) {
            console.log(`Updating timestamp for last approved context: ${lastApprovedContext.prop}`);
            request.activity.updateTimestampEntry(lastApprovedContext.prop);
        }
        request.requestHandler.handleRequestApproved();
    }
}

/**
 * Handles request submission
 * @param {Object} e - The event object
 * @returns {TextOutput} JSON response
 */
function handleRequestSubmission(e) {
    const payload = JSON.parse(e.postData.contents);
    Logger.log("[handleRequestSubmission] Received payload:");
    Logger.log(JSON.stringify(payload, null, 2));

    // Validate payload
    const validation = validateRequestPayload(payload);
    if (!validation.isValid) {
        Logger.log("[handleRequestSubmission] Payload validation failed:");
        Logger.log(validation.error);
        return createErrorResponse(validation.error);
    }

    Logger.log("[handleRequestSubmission] Payload validation passed");

    const validatedPayload = validation.payload;
    const valuesDict = prepareValuesDictionary(validatedPayload);

    // Get appropriate sheet
            const sheetName = getSheetName(valuesDict.REQUEST_TYPE);
    if (!sheetName) {
        return createErrorResponse('Request Type is Invalid.');
    }

    const sheet = getMasterSpreadsheet(sheetName);
    const rowIndex = sheet.getLastRow() + 1;

    // Create event object for onSubmit
    const submitEvent = {
        source: { getActiveSheet: () => sheet },
        range: {
            getRow: () => rowIndex,
            getSheet: () => sheet,
            getCol: () => null,
        },
        valuesDict: valuesDict,
    };

    // Call onSubmit and get the actual final row number
    const finalRowIndex = onSubmit(submitEvent);

    // Use the final row index for processing the request
    const request = new Request(sheet, finalRowIndex);
    processRequestSync(request, validatedPayload);

    // Prepare response data
    const { REQUEST_NUMBER, ATTACHMENT } = request.activity.getActivityValueMap();
    return createSuccessResponse({
        message: "Request submitted successfully",
        requestNumber: REQUEST_NUMBER,
        attachmentUrl: ATTACHMENT,
        timestamp: valuesDict.TIMESTAMP
    });
}

function callMasterApiToUpdateWorkload(mdmName, seconds) {
    if (!mdmName || !seconds || seconds === 0) return;

    // Payload sesuai dengan yang diharapkan oleh handleRequest di Master
    const payload = {
        action: 'update_workload', // Router Master akan menangkap ini
        mdmName: mdmName,
        seconds: seconds
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        Logger.log(`[API] Sending update for ${mdmName} (${seconds}s) to Master...`);
        const response = UrlFetchApp.fetch(WEB_APP_URL, options); // Pastikan konstanta ini ada
        
        const respJson = JSON.parse(response.getContentText());
        if (respJson.status === 'success') {
            Logger.log(`[API] Success. New Total: ${respJson.data.newTotal}`);
        } else {
            Logger.log(`[API] Master returned error: ${respJson.message}`);
        }
    } catch (e) {
        Logger.log(`[API] Failed to connect to Master: ${e.message}`);
    }
}
