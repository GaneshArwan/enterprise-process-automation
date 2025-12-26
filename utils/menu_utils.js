function fixGrantAccess() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No rows selected. Please select one or more rows.");
        return;
    }

    // Function to handle empty string inputs
    const handleEmptyInput = (input) => {
        return input === "" ? null : input.split(",").map(email => email.trim());
    };

    // Batch prompt for all contexts to reduce UI blocking
    const ui = SpreadsheetApp.getUi();
    const approverContexts = ATTACHMENT_SYNC_CONTEXTS.slice(1);
    const additionalEmails = {};

    // Single combined prompt to reduce UI blocking
    if (approverContexts.length > 0) {
        const promptText = approverContexts.map(context =>
            `${context.prop}: (comma-separated emails)`
        ).join('\n');

        const response = ui.prompt(
            'Grant Access - Enter emails for approvers',
            `Enter emails for each approver type (leave empty if none):\n${promptText}`,
            ui.ButtonSet.OK_CANCEL
        );

        if (response.getSelectedButton() === ui.Button.CANCEL) {
            return;
        }

        const lines = response.getResponseText().split('\n');
        approverContexts.forEach((context, index) => {
            const emails = lines[index] ? lines[index].trim() : "";
            additionalEmails[context.prop] = handleEmptyInput(emails);
        });
    }

    // Show progress
    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Processing ${totalRows} requests...`, 'Grant Access', 5);

    let processedCount = 0;
    let errorCount = 0;

    // Process rows in batches to avoid timeouts
    const batchSize = 5;
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                const RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                const attachment = request.activity.getAttachment();

                if (!attachment) {
                    Logger.log(`[fixGrantAccess] No attachment found for row ${rowIndex}`);
                    errorCount++;
                    return;
                }

                // Remove protection and clear data validation when no approvers found
                request.attachment.removeProtection(attachment);
                request.attachment.clearProblematicDataValidation(attachment, additionalEmails);

                // Grant access including approvers' and final approvers' emails
                const grantedEmails = request.attachment.grantAccess(attachment, additionalEmails);

                const { REQUEST_NUMBER } = request.activity.getActivityValueMap();
                Logger.log(`[fixGrantAccess] Granted access to ${grantedEmails.length} emails on Request Number: ${REQUEST_NUMBER}`);
                processedCount++;

            } catch (error) {
                Logger.log(`[fixGrantAccess] Error processing row ${rowIndex}: ${error.message}`);
                errorCount++;
            }
        });

        // Update progress
        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Processed ${Math.min(i + batchSize, totalRows)} of ${totalRows} requests...`,
                'Grant Access',
                2
            );
        }
    }

    // Show completion message
    const message = `Grant Access completed!\nProcessed: ${processedCount}\nErrors: ${errorCount}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete', 3);

    if (errorCount > 0) {
        ui.alert(`Grant Access completed with ${errorCount} errors. Check logs for details.`);
    }
}

function fixSyncToChild() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedVisibleRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No visible rows selected. Please select one or more visible rows.");
        return;
    }

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Syncing ${totalRows} visible requests to child spreadsheets...`, 'Sync to Child', 5);

    let processedCount = 0;
    let errorCount = 0;

    // Process in batches to avoid timeouts
    const batchSize = 10;
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                const RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                const { REQUEST_NUMBER } = request.activity.getActivityValueMap();
                request.activityHandler.copyDataToChild();
                Logger.log(`[fixSyncToChild] Synced data to child spreadsheet on Request Number: ${REQUEST_NUMBER}`);
                processedCount++;
            } catch (error) {
                Logger.log(`[fixSyncToChild] Error processing row ${rowIndex}: ${error.message}`);
                errorCount++;
            }
        });

        // Update progress
        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Synced ${Math.min(i + batchSize, totalRows)} of ${totalRows} visible requests...`,
                'Sync to Child',
                2
            );
        }
    }

    const message = `Sync to Child completed!\nProcessed: ${processedCount}\nErrors: ${errorCount}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete', 3);

    if (errorCount > 0) {
        SpreadsheetApp.getUi().alert(`Sync completed with ${errorCount} errors. Check logs for details.`);
    }
}

function fixSyncToMaster() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedVisibleRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No visible rows selected. Please select one or more visible rows.");
        return;
    }

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Syncing ${totalRows} visible requests to master spreadsheet...`, 'Sync to Master', 5);

    let processedCount = 0;
    let errorCount = 0;

    // Process in batches to avoid timeouts
    const batchSize = 10;
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                const RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                const { REQUEST_NUMBER } = request.activity.getActivityValueMap();

                request.activityHandler.copyDataToMaster();
                Logger.log(`[fixSyncToMaster] Synced data to Master spreadsheet on Request Number: ${REQUEST_NUMBER}`);
                processedCount++;
            } catch (error) {
                Logger.log(`[fixSyncToMaster] Error processing row ${rowIndex}: ${error.message}`);
                errorCount++;
            }
        });

        // Update progress
        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Synced ${Math.min(i + batchSize, totalRows)} of ${totalRows} visible requests...`,
                'Sync to Master',
                2
            );
        }
    }

    const message = `Sync to Master completed!\nProcessed: ${processedCount}\nErrors: ${errorCount}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete', 3);

    if (errorCount > 0) {
        SpreadsheetApp.getUi().alert(`Sync completed with ${errorCount} errors. Check logs for details.`);
    }
}

function fixAttachmentSync() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No rows selected. Please select one or more rows.");
        return;
    }

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Syncing attachment data for ${totalRows} requests...`, 'Attachment Sync', 5);

    let processedCount = 0;
    let errorCount = 0;

    // Process in batches to avoid timeouts
    const batchSize = 5; 
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                let RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                const { REQUEST_NUMBER } = request.activity.getActivityValueMap();

                request.handleOnInterval(REQUEST_NUMBER);
                Logger.log(`[fixAttachmentSync] Attachment Data Synced on Request Number: ${REQUEST_NUMBER}`);
                processedCount++;
            } catch (error) {
                Logger.log(`[fixAttachmentSync] Error processing row ${rowIndex}: ${error.message}`);
                errorCount++;
            }
        });

        // Update progress
        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Synced ${Math.min(i + batchSize, totalRows)} of ${totalRows} requests...`,
                'Attachment Sync',
                2
            );
        }
    }

    const message = `Attachment Sync completed!\nProcessed: ${processedCount}\nErrors: ${errorCount}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete', 3);

    if (errorCount > 0) {
        SpreadsheetApp.getUi().alert(`Attachment Sync completed with ${errorCount} errors. Check logs for details.`);
    }
}

function copyDataToChildAll() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = spreadsheet.getSheets();

    // Filter sheets to only include Activity sheet names
    const validSheets = allSheets.filter(sheet =>
        Object.values(ActivitySheetNames).includes(sheet.getName())
    );

    if (validSheets.length === 0) {
        SpreadsheetApp.getUi().alert("No valid activity sheets found.");
        return;
    }

    // Confirmation dialog
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Copy Data to Child (All)',
        `This will process all rows with TIMESTAMP_ENTRY but empty TAKEN_DATE across ${validSheets.length} activity sheets:\n\n` +
        `• ${validSheets.map(s => s.getName()).join('\n• ')}\n\n` +
        `Continue?`,
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
        return;
    }

    Logger.log(`[copyDataToChildAll] Starting processing of ${validSheets.length} sheets`);

    let totalProcessed = 0;
    let totalErrors = 0;
    let sheetsWithRows = 0;

    // Show initial progress
    spreadsheet.toast(`Processing ${validSheets.length} activity sheets...`, 'Copy Data to Child (All)', 5);

    validSheets.forEach((sheet, sheetIndex) => {
        const sheetName = sheet.getName();
        Logger.log(`[copyDataToChildAll] Processing sheet: ${sheetName} (${sheetIndex + 1}/${validSheets.length})`);

        try {
            // Get rows with timestamp entry but no taken date
            const targetRows = getRowsWithTimestampEntryButNoTakenDate(sheet);

            if (targetRows.length === 0) {
                Logger.log(`[copyDataToChildAll] No matching rows found in sheet: ${sheetName}`);
                return;
            }

            sheetsWithRows++;
            Logger.log(`[copyDataToChildAll] Processing ${targetRows.length} rows in sheet: ${sheetName}`);

            let sheetProcessed = 0;
            let sheetErrors = 0;

            // Process each row in batches
            const batchSize = 5;
            for (let i = 0; i < targetRows.length; i += batchSize) {
                const batchRows = targetRows.slice(i, i + batchSize);

                batchRows.forEach(rowIndex => {
                    try {
                        const headerRow = sheet.getRange(ACTIVITY_HEADER_ROW_INDEX, 1, 1, sheet.getLastColumn()).getValues()[0];
                        const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

                        const rowData = {};
                        headerRow.forEach((header, index) => {
                            if (header) {
                                rowData[header] = rowValues[index];
                            }
                        });

                        let targetSheetName = sheetName;
                        const activitySheetName = getSheetName(rowData[ColNames.REQUEST_TYPE]);
                        if (activitySheetName) {
                            targetSheetName = activitySheetName;
                        } else {
                            sheetErrors++;
                            return;
                        }

                        const RequestClass = getRequestClass(targetSheetName);
                        const request = new RequestClass(sheet, rowIndex, null, rowData);

                        if (!rowData[ColNames.PROCESSED_BY]) {
                            request.requestHandler.handleRequestApproved();
                        }

                        const copyResult = request.activityHandler.copyDataToChild();

                        if (copyResult.success) {
                            sheetProcessed++;
                        } else {
                            sheetErrors++;
                        }

                    } catch (error) {
                        sheetErrors++;
                        Logger.log(`[copyDataToChildAll] Error: ${error.message}`);
                    }
                });

                if (i + batchSize < targetRows.length) {
                    spreadsheet.toast(
                        `Sheet ${sheetName}: ${Math.min(i + batchSize, targetRows.length)}/${targetRows.length} rows processed...`,
                        'Copy Data to Child (All)',
                        2
                    );
                }
            }

            totalProcessed += sheetProcessed;
            totalErrors += sheetErrors;

        } catch (error) {
            totalErrors++;
            Logger.log(`[copyDataToChildAll] Error processing sheet ${sheetName}: ${error.message}`);
        }

        if (sheetIndex + 1 < validSheets.length) {
            spreadsheet.toast(
                `Completed ${sheetIndex + 1}/${validSheets.length} sheets.`,
                'Copy Data to Child (All)',
                2
            );
        }
    });

    const message = `Copy Data to Child (All) completed!\nSheets: ${sheetsWithRows}/${validSheets.length}\nRows: ${totalProcessed}\nErrors: ${totalErrors}`;
    spreadsheet.toast(message, 'Complete', 5);

    if (totalErrors > 0) {
        SpreadsheetApp.getUi().alert(`Completed with ${totalErrors} errors. Check logs.`);
    }
}

function fixOnSubmit() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No rows selected.");
        return;
    }

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Re-triggering onSubmit for ${totalRows} requests...`, 'Fix OnSubmit', 5);

    let processedCount = 0;
    let errorCount = 0;

    const batchSize = 8;
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                let RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                request.handleOnSubmit();
                processedCount++;
            } catch (error) {
                Logger.log(`[fixOnSubmit] Error: ${error.message}`);
                errorCount++;
            }
        });

        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Processed ${Math.min(i + batchSize, totalRows)} of ${totalRows} requests...`,
                'Fix OnSubmit',
                2
            );
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Fix OnSubmit completed! Processed: ${processedCount}`, 'Complete', 3);
}

function fixRequestApproved() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        SpreadsheetApp.getUi().alert("No rows selected.");
        return;
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Fix Request Approved',
        `Re-trigger approval process for ${rowIndexList.length} requests?\n\nThis includes task validation, allocation, protection, emails, and data copying.`,
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Re-triggering approval...`, 'Fix Request Approved', 5);

    let processedCount = 0;
    let errorCount = 0;

    const batchSize = 3;
    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach(rowIndex => {
            try {
                const RequestClass = getRequestClass(sheet.getName());
                const request = new RequestClass(sheet, rowIndex);
                request.requestHandler.handleRequestApproved();
                processedCount++;
            } catch (error) {
                errorCount++;
                Logger.log(`[fixRequestApproved] Error: ${error.message}`);
            }
        });

        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(
                `Processed ${Math.min(i + batchSize, totalRows)} of ${totalRows}...`,
                'Fix Request Approved',
                2
            );
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Completed! Processed: ${processedCount}`, 'Complete', 3);
}

function mergeSelectedVBS() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();
    var mergedContent = "";

    var rowIndex = range.getRow();
    const activity = new Activity(sheet, rowIndex);
    var generator = new ScriptGenerator(activity);
    var sapHeader = generator.initialSAPscript();
    var headerAdded = false;

    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            var fileUrl = values[i][j];
            var fileId = extractSheetId(fileUrl);
            if (fileId) {
                var file = DriveApp.getFileById(fileId);
                var content = file.getBlob().getDataAsString();

                if (content.startsWith(sapHeader)) {
                    content = content.substring(sapHeader.length).trim();
                }

                if (!headerAdded) {
                    mergedContent += sapHeader;
                    headerAdded = true;
                }
                mergedContent += "\n" + content + "\n";
            }
        }
    }

    if (mergedContent) {
        var finalFileUrl = generator.createScriptDoc(mergedContent);
        SpreadsheetApp.getUi().alert("Merged file created:\n" + finalFileUrl);
    } else {
        SpreadsheetApp.getUi().alert("No valid VBS files selected!");
    }
}

function openLink(url) {
    if (!isUrl(url)) {
        SpreadsheetApp.getUi().alert("Invalid URL or configuration.");
        return;
    }
    const html = `<script>window.open("${url}", "_blank");google.script.host.close();</script>`;
    const userInterface = HtmlService.createHtmlOutput(html);
    SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening Link...');
}

function openMDMToolkit() { openLink(MENU_LINK["MDM Toolkit"]); }
function openDashboard() { openLink(MENU_LINK["Dashboard"]); }
function openMDMKnowledgeCenter() { openLink(MENU_LINK["MDM Knowledge Center"]); }
function openScriptMaker() { openLink(SpreadsheetApp.openById(getAttachmentUID("Script Maker")).getUrl()); }
function openMasterProd() { openLink(getMasterSpreadsheet().getUrl()); }

// Replaced specific company names with Generic BU Codes from constants.js
function openRetailA() { openLink(getRequestConfig("BU01")[CHILD_SPREADSHEET_KEY]); }
function openRetailB() { openLink(getRequestConfig("BU02")[CHILD_SPREADSHEET_KEY]); }
function openIndustrial() { openLink(getRequestConfig("BU05")[CHILD_SPREADSHEET_KEY]); }
function openIndustrialB() { openLink(getRequestConfig("BU06")[CHILD_SPREADSHEET_KEY]); }
function openHomeEssentials() { openLink(getRequestConfig("BU03")[CHILD_SPREADSHEET_KEY]); }
function openManufacturing() { openLink(getRequestConfig("BU14")[CHILD_SPREADSHEET_KEY]); }
function openDigital() { openLink(getRequestConfig("BU18")[CHILD_SPREADSHEET_KEY]); }

function changeMDM() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const sourceSheetName = sheet.getName();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        return;
    }

    const title = "Transfer Request to Agent";
    // Sanitized Prompt
    const promptMessage = "Enter target sheet name.\n- Single: AGENT_01\n- Multiple: AGENT_01, AGENT_02, AGENT_03 (PIC is first sheet only)";
    const response = ui.prompt(title, promptMessage, ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
        return;
    }

    const targetSheetNames = response.getResponseText().trim().split(',').map(n => n.trim().toUpperCase());
    const primarySheetName = targetSheetNames[0];
    const primarySheet = spreadsheet.getSheetByName(primarySheetName);

    if (!primarySheet) {
        ui.alert(`Error: Target Sheet "${primarySheetName}" not found.`);
        return;
    }

    const processed = [];
    const failed = [];
    const allSheetNames = targetSheetNames.join(', ');
    const totalRows = rowIndexList.length;

    SpreadsheetApp.getActiveSpreadsheet().toast(`Moving ${totalRows} requests to ${allSheetNames}...`, 'Change MDM', 5);

    rowIndexList.forEach((rowIndex, idx) => {
        try {
            withRowLock(sheet.getName(), rowIndex, 'changeMDM-source', (_srcLock, beatSrc) => {
                const RequestCtor = getRequestClass(sheet.getName());
                const request = new RequestCtor(sheet, rowIndex);

                const { REQUEST_NUMBER, REQUEST_TYPE, DEPARTMENT, ATTACHMENT, COMPANY_CODE_NAME } = request.activity.getActivityValueMap();
                const rawVal = getValueByColumn(sheet, ColNames.ESTIMATED_TIME, rowIndex, ACTIVITY_HEADER_ROW_INDEX);

                let secondsToTransfer = 0;
                
                // Parsing
                const parsedVal = parseFloat(String(rawVal));
                
                // Validasi
                if (Number.isFinite(parsedVal)) {
                    secondsToTransfer = parsedVal;
                } else {
                    Logger.log(`[changeMDM] ❌ Invalid Value! parsedVal=${parsedVal}. Defaulting to 0.`);
                    secondsToTransfer = 0;
                }

                Logger.log(`[changeMDM] ✅ Seconds to Transfer: ${secondsToTransfer}`);
                // =========================================================
                const originalSheetName = sheet.getName();

                if (targetSheetNames.length > 1) {
                    const attachmentId = ATTACHMENT ? extractSheetId(ATTACHMENT) : null;
                    if (attachmentId) addDriveEditors(attachmentId, [EMAIL_MDM_GROUP]);
                }

                withKeyLock(`changeMDM:copy:${primarySheetName}:${REQUEST_NUMBER}`, 'changeMDM-copy', (_k, beatKey) => {
                    request.activityHandler.copyData(primarySheet, ACTIVITY_HEADER_ROW_INDEX, true);
                    beatKey();
                }, 2, 10000);

                const activitySheetName = getSheetName(REQUEST_TYPE);
                const masterSheet = getMasterSpreadsheet(activitySheetName);
                if (masterSheet) {
                    const masterRowIndex = getRowIndex(masterSheet, REQUEST_NUMBER);
                    if (masterRowIndex !== -1) {
                        withRowLock(masterSheet.getName(), masterRowIndex, 'changeMDM-master', (_mLock, beatM) => {
                            const RequestMasterCtor = getRequestClass(masterSheet.getName());
                            const requestMaster = new RequestMasterCtor(masterSheet, masterRowIndex);
                            beatM();
                            requestMaster.activity.updateProcessedBy(allSheetNames);
                        }, 2, 8000);
                    }
                }

                if (secondsToTransfer > 0) {
                    callMasterApiToUpdateWorkload(sourceSheetName, -secondsToTransfer);
                    callMasterApiToUpdateWorkload(primarySheetName, secondsToTransfer);
                }

                // Log Workspace
                logMDMWorkspace(
                    spreadsheet, 
                    REQUEST_NUMBER, 
                    "Change MDM",
                    COMPANY_CODE_NAME, 
                    DEPARTMENT, 
                    REQUEST_TYPE,
                    originalSheetName, 
                    allSheetNames
                );

                processed.push({ req: REQUEST_NUMBER, srcRowAtSelection: rowIndex });

            }, 2, 12000);

            if ((idx + 1) % 5 === 0 && idx + 1 < totalRows) {
                SpreadsheetApp.getActiveSpreadsheet().toast(`Processed ${idx + 1} of ${totalRows}...`, 'Change MDM', 2);
            }
        } catch (e) {
            failed.push({ rowIndex, error: e.message });
        }
    });

    if (processed.length > 0) {
        const rowsToDelete = processed
            .map(p => ({ req: p.req, curIdx: getRowIndex(sheet, p.req) }))
            .filter(x => x.curIdx !== -1)
            .sort((a, b) => b.curIdx - a.curIdx);

        rowsToDelete.forEach(({ curIdx, req }) => {
            try {
                withRowLock(sheet.getName(), curIdx, 'changeMDM-delete', (_dLock) => {
                    sheet.deleteRow(curIdx);
                }, 2, 8000);
            } catch (e) {
                failed.push({ rowIndex: curIdx, error: `Delete failed: ${e.message}` });
            }
        });
    }

    const message = `Change MDM completed!\nMoved: ${processed.length}\nFailed: ${failed.length}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Complete', 3);
    if (failed.length > 0) {
        ui.alert(`Completed with ${failed.length} errors.`);
    }
}

function setDepartmentToSpecialProject() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const rowIndexList = getSelectedRowIndex();

    if (rowIndexList.length === 0) {
        ui.alert("No rows selected.");
        return;
    }

    const response = ui.alert("Confirm", `Update ${rowIndexList.length} requests to 'SPECIAL PROJECT'?`, ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) return;

    const totalRows = rowIndexList.length;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Updating...`, 'Set Special Project', 5);

    const updatedRequests = [];
    const failedRequests = [];
    const batchSize = 10;

    for (let i = 0; i < rowIndexList.length; i += batchSize) {
        const batchRows = rowIndexList.slice(i, i + batchSize);

        batchRows.forEach((rowIndex, batchIndex) => {
            try {
                const request = new (getRequestClass(sheet.getName()))(sheet, rowIndex);
                const { REQUEST_NUMBER, REQUEST_TYPE, DEPARTMENT, COMPANY_CODE_NAME } = request.activity.getActivityValueMap();
                const oldValue = DEPARTMENT; // Simpan departemen lama
                const newValue = "SPECIAL PROJECT";

                request.activity.updateValue("DEPARTMENT", newValue);

                const activitySheetName = getSheetName(REQUEST_TYPE);
                if (!activitySheetName) throw new Error(`Activity sheet not found`);

                const masterSheet = getMasterSpreadsheet(activitySheetName);
                const masterRowIndex = getRowIndex(masterSheet, REQUEST_NUMBER);

                if (masterRowIndex !== -1) {
                    const requestMaster = new (getRequestClass(masterSheet.getName()))(masterSheet, masterRowIndex);
                    requestMaster.activity.updateValue("DEPARTMENT", newValue);

                    logMDMWorkspace(
                        spreadsheet,
                        REQUEST_NUMBER,
                        "Set Department to Special Project",
                        COMPANY_CODE_NAME,
                        oldValue,
                        REQUEST_TYPE,
                        oldValue,
                        newValue
                    );

                    updatedRequests.push(REQUEST_NUMBER);
                } else {
                    throw new Error(`Master row not found`);
                }
            } catch (e) {
                failedRequests.push(`Row ${rowIndex}: ${e.message}`);
            }
        });

        if (i + batchSize < rowIndexList.length) {
            SpreadsheetApp.getActiveSpreadsheet().toast(`Updated ${Math.min(i + batchSize, totalRows)}...`, 'Set Special Project', 2);
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Completed! Updated: ${updatedRequests.length}`, 'Complete', 3);
    if (failedRequests.length > 0) ui.alert(`Failed: ${failedRequests.length} rows.`);
}

function runPrioritySorting() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();

    if (Object.values(MDMSheetNames).includes(sheetName)) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Sorting '${sheetName}'...`, 'Processing', 5);
        prioritySorting(sheet);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Sorted.`, 'Success', 3);
    } else {
        SpreadsheetApp.getUi().alert(`This function can only be run on Agent Worksheets.`);
    }
}