/**
 * Represents an attachment within a spreadsheet.
 * 
 * @param {Sheet} sheet The Google Sheets sheet where the attachment is located.
 * @param {number} rowIndex The row index of the attachment within the sheet.
 */
class Attachment {
    constructor(sheet, rowIndex) {
        this.sheet = sheet;
        this.rowIndex = rowIndex;
        this.activity = new ActivityWithUpdate(
            sheet, rowIndex);
        this.email = new EmailHandler(
            sheet, rowIndex
        );
        this.attachment = null;
    }

    /**
     * Retrieves the attachment spreadsheet.
     * 
     * @returns {Spreadsheet} The attachment spreadsheet.
     */
    getAttachment() {
        if (this.attachment == null) {
            this.attachment = this.activity.getAttachment();
        }

        return this.attachment;
    }

    /**
     * Generates a file name for the attachment based on activity values.
     * 
     * @returns {string} The generated file name.
     */
    getFileName() {
        let {
            REQUEST_NUMBER, REQUEST_TYPE,
            DEPARTMENT, EMAIL_ADDRESS
        } = this.activity.getActivityValueMap();

        DEPARTMENT = DEPARTMENT ? DEPARTMENT : ''

        return `${REQUEST_NUMBER} ${REQUEST_TYPE} ${DEPARTMENT} ${EMAIL_ADDRESS} ${getDateNow()}`;
    }

    getValuesByCell(cells) {
        const attachment = this.getAttachment();
        return cells.map(cell =>
            attachment.getRange(cell).getValue()
        );
    }

    updateValuesByCell(cells, values) {
        const attachment = this.getAttachment();
        
        if (cells.length === 1) {
            // Single cell update
            attachment.getRange(cells[0]).setValue(values[0]);
        } else if (cells.length > 1) {
            // Check if we can batch update contiguous ranges
            try {
                // Try to create a single range if cells are contiguous
                const ranges = cells.map(cell => attachment.getRange(cell));
                
                // For now, use individual updates but log for potential future optimization
                cells.forEach((cell, index) => {
                    attachment.getRange(cell).setValue(values[index]);
                });
                
                Logger.log(`[updateValuesByCell] Updated ${cells.length} cells in attachment`);
            } catch (error) {
                Logger.log(`[updateValuesByCell] Error updating cells: ${error.message}`);
                // Fallback to individual updates
                cells.forEach((cell, index) => {
                    try {
                        attachment.getRange(cell).setValue(values[index]);
                    } catch (e) {
                        Logger.log(`[updateValuesByCell] Failed to update cell ${cell}: ${e.message}`);
                    }
                });
            }
        }
    }

    insertValues(targetSheet, headers, values) {
        const headersRowIndex = AttachmentValues.TASK_START_ROW - 1;
        const targetHeaders = getColumnHeaders(targetSheet, true, headersRowIndex);
        if (!targetHeaders) return;

        headers = headers.map(header => upperAndSeparate(header));

        // Map values to targetHeaders order for each row in values
        const mappedValues = values.map(row =>
            targetHeaders.map(targetHeader => {
                // Find the index of the matching header
                const headerIndex = headers.findIndex(header => header === targetHeader);

                // If a match is found, get the corresponding value from the row,
                // otherwise use an empty string
                return headerIndex !== -1 ? row[headerIndex] : '';
            })
        );

        const startRow = AttachmentValues.TASK_START_ROW;
        const startColumn = 1;
        targetSheet
            .getRange(startRow, startColumn, mappedValues.length, targetHeaders.length)
            .setValues(mappedValues);
    }

    /**
     * Retrieves requester values from the attachment.
     * 
     * @returns {Object} An object containing requester name and status.
     */
    getRequesterValues() {
        const attachment = this.getAttachment();
        return {
            name: attachment.getRange(
                AttachmentValues.REQUESTER_NAME_CELL
            ).getValue(),

            status: attachment.getRange(
                AttachmentValues.REQUESTER_STATUS_CELL
            ).getValue()
        }
    }

    /**
     * Calculates the total number of tasks in the attachment.
     * 
     * @returns {number} The total number of tasks.
     */
    getTotalTask() {
        const attachment = this.getAttachment();
        let totalRows = 0;

        if (!attachment) return 1;

        const isPromotionCreate = AttachmentValues.TASK_START_ROW === 34;

        attachment.getSheets().forEach(sheet => {
            if (sheet.getTabColor() === TASK_SHEET_COLOR) {
                const lastRow = sheet.getLastRow();

                // If no meaningful task rows, skip
                if (lastRow < AttachmentValues.TASK_START_ROW) return;

                const lastColumn = isPromotionCreate
                    ? sheet.getLastColumn() - 3 // Exclude the last 3 columns for Promotion Create
                    : sheet.getLastColumn();

                // Get the range of data rows
                const range = sheet.getRange(
                    AttachmentValues.TASK_START_ROW,
                    1,
                    lastRow - AttachmentValues.TASK_START_ROW + 1,
                    lastColumn
                );

                const values = range.getValues();

                // Check if there are no non-empty rows in the data range
                const nonEmptyRows = values.filter(row =>
                    row.some(cell => cell !== '' && cell !== null && cell !== undefined)
                ).length;

                if (nonEmptyRows === 0) return;  // Skip if no non-empty rows found

                totalRows += nonEmptyRows;
            }
        });

        return totalRows;
    }


    /**
     * Retrieves the request template based on the activity type.
     * 
     * @returns {File} The request template file.
     */
    getRequestTemplate(templateUrl) {
        let attachmentUID = templateUrl
            ? extractSheetId(templateUrl)
            : getAttachmentUID(this.activity.getActivityValueMap().REQUEST_TYPE);

        return DriveApp.getFileById(attachmentUID);
    }

    setApproverNote() {
        const attachment = this.getAttachment();
        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const ctx = ATTACHMENT_SYNC_CONTEXTS[i];
            const { prop } = ctx;
            if (!this.activity.hasCol(`RESPON_${prop}`)) return;

            let approverEmails = [];
            const isApprover = this.email.isEmailApprover(ctx);
            if (isApprover) {
                approverEmails = this.email.getEmailApprover(ctx);
            }
            
            const rangeNote = attachment.getRange(AttachmentValues[`${prop}_NAME_CELL`]);
            
            if (approverEmails.length > 0) {
                // Only set dropdown validation if there are actual approvers
                this.setDropdownWithOptions(
                    attachment,
                    AttachmentValues[`${prop}_NAME_CELL`],
                    approverEmails,
                    false,
                    false
                );
                rangeNote.setNote(`List ${prop} :\n${approverEmails.join("\n")}`);
            } else {
                // Clear any existing data validation and set note for auto-approval
                rangeNote.clearDataValidations();
                rangeNote.setNote(`List ${prop} : Tidak ada Approver (Auto Approved)`);
            }
            Logger.log(`[setApproverNote] Setting notes approver emails for ${prop}: ${approverEmails.length} emails`);
        }
    }

    grantAccess(attachment, additionalEmails = {}, isProtect = true) {
        let emails = this.email.getAllEmails();

        Object.keys(additionalEmails).forEach(key => {
            const emailList = additionalEmails[key];
            if (Array.isArray(emailList)) {
                emails.push(...emailList.filter(email => email && email.trim()));
            } else if (emailList && emailList.trim()) {
                emails.push(emailList.trim());
            }
        });

        // Remove duplicates and filter out empty emails
        emails = [...new Set(emails.filter(email => email && email.trim()))];

        try {
            addDriveEditors(attachment.getId(), emails);
            Logger.log(`[grantAccess] Added ${emails.length} editors to attachment`);
        } catch (error) {
            Logger.log(`[grantAccess] Error adding editors: ${error.message}`);
        }

        // Set approver notes with proper validation handling
        this.setApproverNote();

        if (isProtect) {
            this.protectApprover(attachment, additionalEmails);
        }

        return emails;
    }
    /**
     * Creates a copy of the attachment with optional image folder.
     * 
     * @param {boolean} withImageFolder Indicates if an image folder should be created.
     * @param {Range} imageCell The cell containing the image URL.
     * @returns {Object} An object containing the attachment and image folder URL.
     */
    makeAttachmentCopy(withImageFolder, imageCell, templateUrl = null, attachmentValues = null) {
        const { requestDrive, imageDrive } = getDriveFolder(
            this.activity.getBaseName(),
            this.activity.getCompanyName(),
            withImageFolder
        );

        if (!requestDrive) return;

        const fileName = this.getFileName();
        const template = this.getRequestTemplate(templateUrl);

        let attachment = template.makeCopy(fileName, requestDrive);
        attachment = SpreadsheetApp.openByUrl(attachment.getUrl());
        this.attachment = attachment;
        console.log("this attachment created: ", this.attachment.getName())
        this.grantAccess(attachment); //Grant Approver Access
        if (templateUrl) {
            this.removeProtection(attachment); //Remove Protection if template is used
        }

        if (attachmentValues) {
            for (const [key, value] of Object.entries(attachmentValues)) {
                const targetSheet = attachment.getSheets()[key]
                const headers = value[0]
                const values = value.slice(1);

                this.insertValues(targetSheet, headers, values)
            }
        }

        let imageFolderURL = null;
        if (withImageFolder && imageCell) {
            const { REQUEST_NUMBER } = this.activity.getActivityValueMap();
            const imageFolder = imageDrive.createFolder(REQUEST_NUMBER)
            this.grantAccess(imageFolder, { MDM: EMAIL_MDM_GROUP }, false)
            this.setImageFolderUrl(imageFolder.getUrl(), imageCell)
        }

        this.setDefaultValues();

        return {
            attachment: attachment,
            imageFolderURL: imageFolderURL
        };
    }

    /**
     * Retrieves values from a specific sheet within the attachment.
     * 
     * @param {string} sheetName The name of the sheet to retrieve values from.
     * @param {number} [startRow=1] The starting row for value retrieval.
     * @returns {Array} A 2D array of values from the specified sheet.
     */
    getValuesBySheet(sheetName, startRow = 2) {
        const attachment = this.getAttachment();
        if (!attachment) return;

        const sheet = attachment.getSheets().find(s =>
            s.getName().includes(sheetName)
        );
        if (!sheet) return;

        const lastRow = getLastNonEmptyRow(sheet);

        if (lastRow < startRow) return
        return getRowValuesMap(sheet, startRow, lastRow, true, startRow - 1);
    }

    /**
     * Retrieves feedback values from the attachment.
     * 
     * @returns {Array} A 2D array of feedback values.
     */
    getFeedbackValues() {
        // Retrieve and return values from the feedback sheet
        const feedbackValues = this.getValuesBySheet(SheetNames.MDM_FEEDBACK);
        if (!feedbackValues) return;

        return feedbackValues
    }

    /**
     * Sets default values for the attachment.
     */
    setDefaultValues() {
        Logger.log(`[setDefaultValues] Setting default values for attachment`);
        const attachment = this.getAttachment();
        const companyName = this.activity.getCompanyName();
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap();
        
        console.log("EMAIL ADDRESS: ", EMAIL_ADDRESS)
        // Prepare batch update data
        const updates = [
            { range: AttachmentValues.COMPANY_CELL, value: getCompanyFullName(companyName) },
            { range: AttachmentValues.REQUESTER_NAME_CELL, value: EMAIL_ADDRESS }
        ];
        
        // Try to optimize with batch operation if possible
        try {
            // Use batch API if available, otherwise fall back to individual updates
            const ranges = updates.map(u => u.range);
            const values = updates.map(u => u.value);
            
            // Call optimized batch update method
            this.updateValuesByCell(ranges, values);
            
            Logger.log(`[setDefaultValues] Successfully set ${updates.length} default values in batch`);
        } catch (error) {
            Logger.log(`[setDefaultValues] Batch update failed, using individual updates: ${error.message}`);
            // Fallback to individual updates
            updates.forEach(({range, value}) => {
                attachment.getRange(range).setValue(value);
            });
        }
    }

    setDropdownWithOptions(sheet, cellsRef, optionsArray, invalid = false, showList = false) {
        const range = sheet.getRange(cellsRef);
        const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(optionsArray, showList)
            .setAllowInvalid(invalid)
            .build();
        range.setDataValidation(rule);
    }

    clearValuesByCell(cells) {
        const attachment = this.getAttachment();
        cells.forEach((cell, _) => {
            attachment.getRange(cell).clearContent();
        });
    }

    clearApproverValues(ctx) {
        const { prop, status, name } = ctx;

        if (this.hasApproverStatus(ctx)) {
            this.clearValuesByCell(
                [
                    AttachmentValues[`${prop}_NAME_CELL`],
                    AttachmentValues[`${prop}_STATUS_CELL`]
                ],
                [name, status]
            )
        }
    }


    /**
     * Clears requester values from the attachment.
     */
    clearRequesterValues() {
        Logger.log(`[clearRequesterValues] Clearing requester values from attachment`);
        this.getAttachment().getRange(AttachmentValues.REQUESTER_STATUS_CELL).clearContent();
        Logger.log(`[clearRequesterValues] Requester values cleared successfully`);
    }

    /**
     * Handles expired attachments by renaming and protecting them.
     */
    handleAttachmentExpired() {
        Logger.log(`[handleAttachmentExpired] Handling expired attachment`);
        this.getAttachment().rename(RequesterStatus.EXPIRED + " " + this.getFileName());
        this.protectSpreadsheet();
        Logger.log(`[handleAttachmentExpired] Attachment marked as expired and protected`);
    }

    updateRequesterValues(status, name) {
        setValueWithCellRef(
            this.getAttachment(), AttachmentValues.REQUESTER_STATUS_CELL,
            status
        )

        setValueWithCellRef(
            this.getAttachment(), AttachmentValues.REQUESTER_NAME_CELL,
            name
        )
    }

    updateApproverValues(ctx) {
        const { prop, status, name } = ctx;

        if (this.hasApproverStatus(ctx)) {
            this.updateValuesByCell(
                [
                    AttachmentValues[`${prop}_NAME_CELL`],
                    AttachmentValues[`${prop}_STATUS_CELL`]
                ],
                [name, status]
            )
        }
    }

    handleAttachmentCompleted() {
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap();
        this.updateRequesterValues(
            RequesterStatus.COMPLETED,
            EMAIL_ADDRESS
        )
    }

    /**
     * Protects the attachment spreadsheet by setting permissions.
     */
    protectSpreadsheet() {
        try {
            const attachment = this.getAttachment();
            if (!attachment) {
                Logger.log(`[protectSpreadsheet] No attachment found for row ${this.rowIndex}. Skipping protection.`);
                return;
            }

            // try {
            //     addDriveEditors(attachment.getId(), [EMAIL_MDM_GROUP]);
            //     Logger.log(`[protectSpreadsheet] Successfully granted file access to MDM Group for attachment: ${attachment.getName()}`);
            // } catch (e) {
            //     Logger.log(`[protectSpreadsheet] Failed to grant file access to MDM Group: ${e.message}`);
            // }
            const sheets = attachment.getSheets();
            const userEmail = Session.getEffectiveUser().getEmail(); // Get the executing user's email

            sheets.forEach(sheet => {
                try {
                    const protectedSheet = sheet.protect();
                    const editors = protectedSheet.getEditors();

                    // Remove all editors except the current user
                    editors.forEach(editor => {
                        try {
                            if (editor.getEmail() !== userEmail) {
                                protectedSheet.removeEditor(editor);
                            }
                        } catch (e) {
                            console.warn(`Failed to remove editor: ${e.message}`);
                        }
                    });

                    // Ensure the MDM group is added as an editor
                    protectedSheet.addEditor(EMAIL_MDM_GROUP);
                } catch (e) {
                    console.error(`Error protecting sheet: ${e.message}`);
                }
            });

        } catch (e) {
            console.error(`Error in protectSpreadsheet: ${e.message}`);
        }
    }


    removeProtection(attachment = null) {
        const sheets = attachment?.getSheets() || this.getAttachment().getSheets();
        if (!sheets || sheets.length === 0) return;

        // Remove sheet protection from all sheets
        sheets.forEach(sheet => {
            const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
            sheetProtections.forEach(protection => protection.remove());
        });

        // Only remove range protections from the first sheet
        const firstSheet = sheets[0];
        const rangeProtections = firstSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        rangeProtections.forEach(protection => protection.remove());
    }

    /**
     * Clears data validation from approver name cells when no approvers are found
     * @param {Spreadsheet} attachment - The attachment spreadsheet
     * @param {Object} additionalEmails - Additional emails provided by user
     */
    clearProblematicDataValidation(attachment = null, additionalEmails = {}) {
        const sheet = attachment?.getSheets()?.[0] || this.getAttachment().getSheets()[0];
        if (!sheet) {
            Logger.log('No sheet found from attachment or fallback attachment.');
            return;
        }

        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const ctx = ATTACHMENT_SYNC_CONTEXTS[i];
            const { prop } = ctx;
            
            // Skip if this column doesn't exist in the current activity
            if (!this.activity.hasCol(`RESPON_${prop}`)) continue;

            // Check if there are any approver emails (from system or additional)
            const systemEmails = this.email.getEmailApprover(ctx);
            const additionalPropEmails = additionalEmails[prop] || [];
            const totalEmails = [...systemEmails, ...additionalPropEmails].filter(email => email && email.trim());

            // If no emails, clear data validation from name cell
            if (totalEmails.length === 0) {
                try {
                    const nameCell = AttachmentValues[`${prop}_NAME_CELL`];
                    
                    if (nameCell) {
                        const nameRange = sheet.getRange(nameCell);
                        nameRange.clearDataValidations();
                        nameRange.setNote(`List ${prop} : Tidak ada Approver (Auto Approved)`);
                        Logger.log(`[clearProblematicDataValidation] Cleared data validation for ${prop} name cell - no approvers found`);
                    }
                } catch (error) {
                    Logger.log(`[clearProblematicDataValidation] Error clearing validation for ${prop}: ${error.message}`);
                }
            }
        }
    }

    protectApprover(attachment = null, additionalEmails = {}) {
        const sheet = attachment?.getSheets()?.[0] || this.getAttachment().getSheets()[0];
        if (!sheet) {
            Logger.log('No sheet found from attachment or fallback attachment.');
            return;
        }

        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const ctx = ATTACHMENT_SYNC_CONTEXTS[i];
            const { prop } = ctx;
            if (!this.activity.hasCol(`RESPON_${prop}`)) continue;

            const systemEmails = this.email.getEmailApprover(ctx);
            const additionalPropEmails = additionalEmails[prop] ? additionalEmails[prop] : [];
            const approverEmails = [...systemEmails, ...additionalPropEmails].filter(email => email && email.trim());

            // Only protect ranges if there are actual approvers
            if (approverEmails.length > 0) {
                try {
                    const protectedRange = sheet.getRange(
                        `${AttachmentValues[`${prop}_STATUS_CELL`]}:${AttachmentValues[`${prop}_NOTES_CELL`]}`
                    ).protect();

                    protectedRange.removeEditors(protectedRange.getEditors());
                    let sentEmail = [];
                    
                    approverEmails.forEach(email => {
                        try {
                            protectedRange.addEditor(email.trim());
                            if (!sentEmail.includes(email.trim())) {
                                sentEmail.push(email.trim());
                            }
                        } catch (error) {
                            Logger.log(`Failed to add editor: ${email} - ${error.message}`);
                        }
                    });

                    Logger.log(`${prop} access granted to: ` + sentEmail.join(', '));
                } catch (error) {
                    Logger.log(`[protectApprover] Error protecting range for ${prop}: ${error.message}`);
                }
            } else {
                Logger.log(`${prop} has no approvers - skipping protection`);
            }
        }
    }

    setImageFolderUrl(url, imageCell) {
        const imageSheet = this.getAttachment().getSheetByName(SheetNames.IMAGE);
        setValueWithCellRef(
            imageSheet, imageCell, url
        )

        const imageRange = imageSheet.getRange(imageCell);
        const protection = imageRange.protect();
        protection.removeEditors(protection.getEditors());
    }

}

class AttachmentWithValidation extends Attachment {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex);
    }

    hasApproverStatus(ctx) {
        const { prop } = ctx;
        const cell = AttachmentValues[`${prop}_STATUS_CELL`];
        const statusCell = this.getAttachment()
            .getSheets()[0]
            .getRange(cell);
        const validation = statusCell.getDataValidation();

        return (
            validation != null &&
            validation.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
        );
    }

    //use this to set additional values in the 'ATTACHMENTS' sheet
    setAdditionalAttachmentValues(values) {
        const attachment = this.getAttachment();
        const attachmentSheet = attachment.getSheetByName('ATTACHMENTS');
        if (!attachmentSheet) return;

        setValueWithCellRef(
            attachmentSheet, AttachmentValues.IMAGE_CELL, values
        )
        Logger.log('[Attachment] Additional Attachment added to Attachment.');
    }

    /**
     * Clear all approver status data validations to prevent validation errors
     * Use this when injecting new requests or when approver configurations change
     */
    clearAllApproverStatusValidations() {
        const attachment = this.getAttachment();
        const sheet = attachment.getSheets()[0];
        if (!sheet) return;

        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const ctx = ATTACHMENT_SYNC_CONTEXTS[i];
            const { prop } = ctx;
            
            // Skip if this column doesn't exist in the current activity
            if (!this.activity.hasCol(`RESPON_${prop}`)) continue;

            try {
                const statusCell = AttachmentValues[`${prop}_STATUS_CELL`];
                const nameCell = AttachmentValues[`${prop}_NAME_CELL`];
                
                if (statusCell) {
                    const statusRange = sheet.getRange(statusCell);
                    statusRange.clearDataValidations();
                    statusRange.clearContent(); // Also clear any invalid content
                    Logger.log(`[clearAllApproverStatusValidations] Cleared status validation for ${prop}`);
                }
                
                if (nameCell) {
                    const nameRange = sheet.getRange(nameCell);
                    nameRange.clearDataValidations();
                    Logger.log(`[clearAllApproverStatusValidations] Cleared name validation for ${prop}`);
                }
            } catch (error) {
                Logger.log(`[clearAllApproverStatusValidations] Error clearing validation for ${prop}: ${error.message}`);
            }
        }
        
        // After clearing validations, re-setup proper approver notes
        this.setApproverNote();
    }

}

class AttachmentPromotion extends AttachmentWithValidation {
    constructor(sheet, rowIndex) {
        console.log("Creating AttachmentPromotion instance");
        super(sheet, rowIndex)
    }

    /**
     * Retrieves the request template based on the activity type, with custom promo code handling.
     * 
     * @returns {File} The request template file.
     */
    getRequestTemplate(templateUrl = null) {
        const { REQUEST_TYPE } = this.activity.getActivityValueMap();
        if (REQUEST_TYPE !== RequestTypes.PROMOTION_CREATE || templateUrl) {
            console.log("Using template URL: " + templateUrl);
            return super.getRequestTemplate(templateUrl);
        }

        const promoCode = this.activity.getPromoCode();
        console.log("PROMO CODE: " + promoCode);
        const attachmentUID = getAttachmentUID(promoCode);

        return DriveApp.getFileById(attachmentUID);
    }

    makeAttachmentCopy(withImageFolder, imageCell, templateUrl = null, attachmentValues = null) {
        const { attachment, imageFolderURL } = super.makeAttachmentCopy(
            withImageFolder, imageCell,
            templateUrl, attachmentValues
        );

        console.log("TEMPATE URL MAKE COPY: ", templateUrl);
        const promoCode = this.activity.getPromoCode();
        if (promoCode === "Z005") {
            attachment.addEditor(EMAIL_CC_SECOND)
        }

        return {
            attachment, imageFolderURL
        }
    }

    getTotalTask() {
        const { REQUEST_TYPE } = this.activity.getActivityValueMap();
        if (REQUEST_TYPE === RequestTypes.PROMOTION_CREATE) {
            AttachmentValues = Object.freeze({
                ...AttachmentValues,
                TASK_START_ROW: 34
            });
        }

        return super.getTotalTask();
    }

}

class AttachmentFinanceMaster extends AttachmentWithValidation {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex)
    }

    setLinkBlockUrl(value) {
        const linkBlockSheet = this.getAttachment().getSheetByName(SheetNames.LINK_BLOCK);
        setValueWithCellRef(
            linkBlockSheet, AttachmentValues.IMAGE_CELL, value
        )
        Logger.log('[Attachment] Block Cost Center URL added to attachment');
    }
}

class AttachmentPricing extends AttachmentWithValidation {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex)
    }

    _getCustomAttachmentName() {
        const { PRICING_TYPE } = this.activity.getActivityValueMap();
        const accessSequence = this.activity.getAccessSequence();
        return `${PRICING_TYPE} - ${accessSequence}`
    }

    getRequestTemplate(templateUrl = null) {
        if (templateUrl) {
            return super.getRequestTemplate(templateUrl);
        }

        const attachmentName = this._getCustomAttachmentName();
        Logger.log("Creating Attachment with Name: " + attachmentName);
        const attachmentUID = getAttachmentUID(attachmentName);
        Logger.log("Attachment UID: " + attachmentUID)

        return DriveApp.getFileById(attachmentUID);
    }
}

class AttachmentVendor extends AttachmentWithValidation {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex)
        this.email = new EmailHandlerVendor(sheet, rowIndex);
    }
}