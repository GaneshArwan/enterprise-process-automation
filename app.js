function onSubmit(e) {
    const sheet = e.source.getActiveSheet();
    let sourceRow = e.range.getRow();

    // Only clear stale data cache; keep header metadata to avoid expensive reload
    clearSheetCache(sheet, { skipHeaders: true });

    const RequestClass = getRequestClass(sheet.getName());
    const initialRowData = e.valuesDict || null;
    const request = new RequestClass(sheet, sourceRow, null, initialRowData);

    if (initialRowData) {
        request.insertValues(initialRowData);
    }

    request.handleOnSubmit();
    return sourceRow;
}

/**
 * Re-run onSubmit for rows that previously errored, across all
 * Activity sheets' *_SUBMIT tabs.
 */
function onIntervalFixSubmit() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = Object.values(ActivitySheetNames);

    for (const baseName of sheetNames) {
        const submitSheetName = `${baseName}${SUBMIT_SUFFIX}`;
        const sheet = ss.getSheetByName(submitSheetName);
        if (!sheet) {
            Logger.log(`[onIntervalFixSubmit] Skipping missing sheet: ${submitSheetName}`);
            continue;
        }

        const RequestClass = getRequestClass(baseName);
        const onSubmitErrorRows = getOnSubmitErrorRow(sheet) || [];
        if (onSubmitErrorRows.length === 0) continue;

        Logger.log(
            `[onIntervalFixSubmit] Re-trigger onSubmit ${submitSheetName} rows: ${onSubmitErrorRows.join(', ')}`
        );

        [...onSubmitErrorRows].sort((a, b) => b - a).forEach((row) => {
            if (row <= sheet.getLastRow()) {
                const request = new RequestClass(sheet, row);
                request.handleOnSubmit();
            }
        });
    }
}

function onInterval(sheetNames = []) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    for (let sheet of sheets) {
        const sheetName = sheet.getName();
        const validSheetNames = sheetNames.length > 0 ? sheetNames : Object.values(ActivitySheetNames);
        if (!validSheetNames.includes(sheetName)) {
            continue;
        }

        let RequestClass = getRequestClass(sheet.getName());
        const targetRows = getActivityRowToCrawl(sheet);

        Logger.log(`[onInterval] Processing ${sheetName} on rows: ${targetRows.map(({ row }) => row).join(', ')}`);

        const sortedRows = [...targetRows].sort((a, b) => b.row - a.row);

        // Process rows
        sortedRows.forEach(({ row, requestNumber }) => {
            if (row <= sheet.getLastRow()) {
                const request = new RequestClass(sheet, row);
                const status = request.handleOnInterval(requestNumber);
                if (!status) return;
            }
        });

        const errorSentBackRows = getSystemSentBackErrorRow(sheet);
        if (errorSentBackRows.length > 0) {
            Logger.log(`[onInterval] Fixing System Sent Back Error on ${sheetName} rows: ${errorSentBackRows.join(', ')}`);

            [...errorSentBackRows].sort((a, b) => b - a).forEach(row => {
                if (row <= sheet.getLastRow()) {
                    const request = new RequestClass(sheet, row);
                    request.requestHandler.handleSystemSentBackEmail();
                }
            });
        }
    }
}

function onIntervalExtendPIR() { return onInterval([ActivitySheetNames.EXTEND_PIR]) }
function onIntervalCustomer() { return onInterval([ActivitySheetNames.CUSTOMER]) }
function onIntervalPricing() { return onInterval([ActivitySheetNames.PRICING]) }
function onIntervalMasterSite() { return onInterval([ActivitySheetNames.MASTER_SITE]) }
function onIntervalStatusListing() { return onInterval([ActivitySheetNames.STATUS_LISTING]) }
function onIntervalProfitCenter() { return onInterval([ActivitySheetNames.PROFIT_CENTER]) }
function onIntervalMasterFinance() { return onInterval([ActivitySheetNames.MASTER_FINANCE]) }
function onIntervalHierarchy() { return onInterval([ActivitySheetNames.HIERARCHY]) }
function onIntervalBasicData() { return onInterval([ActivitySheetNames.BASIC_DATA]) }
function onIntervalNonM() { return onInterval([ActivitySheetNames.NON_M]) }
function onIntervalImage() { return onInterval([ActivitySheetNames.IMAGE]) }
function onIntervalBom() { return onInterval([ActivitySheetNames.BOM]) }
function onIntervalPromotion() { return onInterval([ActivitySheetNames.PROMOTION]) }
function onIntervalMasterData() { return onInterval([ActivitySheetNames.MASTER_DATA]) }
function onIntervalSourceList() { return onInterval([ActivitySheetNames.SOURCE_LIST]) }
function onIntervalMerchandise() { return onInterval([ActivitySheetNames.MERCHANDISE]) }
function onIntervalVendor() { return onInterval([ActivitySheetNames.VENDOR]) }

function onChildEdit(e) {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    if (e.range.getRow() <= ACTIVITY_HEADER_ROW_INDEX) return;

    const range = e.range;
    const rowIndex = range.getRow();
    const colIndex = range.getColumn();
    const userEmail = e.user && e.user.getEmail ? e.user.getEmail() : "";
    let oldValueForHandler = e.oldValue ?? null;
    let mutated = false;

    const flushPendingEdits = (reason) => {
        try {
            const flushStart = new Date();
            Logger.log(`[onChildEdit] Flushing pending edits (${reason}) on row ${rowIndex}`);
            SpreadsheetApp.flush();
            Logger.log(`[onChildEdit] Flush completed in ${new Date() - flushStart}ms on row ${rowIndex}`);
        } catch (flushError) {
            Logger.log(`[onChildEdit] Flush warning (${reason}) on row ${rowIndex}: ${flushError.message}`);
        }
    };

    const headerRow = getColumnHeaders(sheet, false, ACTIVITY_HEADER_ROW_INDEX);
    const lastUsedColumn = headerRow.length || sheet.getLastColumn();
    const editedRowValues = sheet.getRange(rowIndex, 1, 1, lastUsedColumn).getValues()[0];

    const rowData = {};
    headerRow.forEach((header, index) => {
        if (header) rowData[header] = editedRowValues[index];
    });

    const processStatusColIdx = getColumnIndex(sheet, ColNames.PROCESS_STATUS, ACTIVITY_HEADER_ROW_INDEX);

    if (colIndex === processStatusColIdx && Object.values(MDMSheetNames).includes(sheetName)) {
        const oldStatus = ((e.oldValue ?? "") + "").trim();
        const newStatus = ((e.value ?? "") + "").trim();
        if (oldStatus === newStatus) return;
        rowData[ColNames.PROCESS_STATUS] = newStatus;

        const takenDateColIdx = getColumnIndex(sheet, ColNames.TAKEN_DATE, ACTIVITY_HEADER_ROW_INDEX);
        const estTimeColIdx = getColumnIndex(sheet, ColNames.ESTIMATED_TIME_FINISHED, ACTIVITY_HEADER_ROW_INDEX);
        const feedbackStatusColIdx = getColumnIndex(sheet, ColNames.FEEDBACK_STATUS, ACTIVITY_HEADER_ROW_INDEX);
        const processedDateColIdx = getColumnIndex(sheet, ColNames.PROCESSED_DATE, ACTIVITY_HEADER_ROW_INDEX);

        const takenDateVal = rowData[ColNames.TAKEN_DATE];
        const processedDateVal = rowData[ColNames.PROCESSED_DATE];

        if (oldStatus === MDMStatus.SEND_BACK && newStatus !== MDMStatus.SEND_BACK) {
            e.source.toast("Status '" + oldStatus + "' cannot be changed or cleared. Reverting.", "Not Allowed", 5);
            e.oldValue = MDMStatus.SEND_BACK;
            range.setValue(MDMStatus.SEND_BACK);
            mutated = true;
            flushPendingEdits('reverting SEND_BACK change');
            return;
        }

        const TERMINAL_STATUSES = new Set([
            MDMStatus.COMPLETED,
            MDMStatus.PARTIALLY_REJECTED,
            MDMStatus.REJECTED,
        ]);
        if (newStatus === MDMStatus.ON_GOING && (TERMINAL_STATUSES.has(oldStatus) || processedDateVal)) {
            e.source.toast("Cannot change '" + oldStatus + "' back to 'On Going'. Reverting.", "Not Allowed", 6);
            range.setValue(oldStatus);
            return;
        }

        if (!newStatus) {
            const colsToClear = [];
            if (takenDateColIdx > 0 && rowData[ColNames.TAKEN_DATE]) colsToClear.push(ColNames.TAKEN_DATE);
            if (estTimeColIdx > 0) colsToClear.push(ColNames.ESTIMATED_TIME_FINISHED);
            if (feedbackStatusColIdx > 0) colsToClear.push(ColNames.FEEDBACK_STATUS);
            if (processedDateColIdx > 0) colsToClear.push(ColNames.PROCESSED_DATE);

            if (colsToClear.length) {
                setValuesWithIndexes(sheet, colsToClear, rowIndex, new Array(colsToClear.length).fill(""));
                colsToClear.forEach(colName => { rowData[colName] = ""; });
                mutated = true;
            }

            Logger.log(`[onChildEdit] Cleared dependent fields on row ${rowIndex}`);
            flushPendingEdits('clearing status-dependent values');
            return;
        }

        if (newStatus !== MDMStatus.ON_GOING && !takenDateVal) {
            e.source.toast("You can only set status to "+ newStatus +" when 'Taken Date' is NOT empty.", "Not Allowed", 6);
            range.setValue(oldStatus);
            mutated = true;
            flushPendingEdits('prevent revert to ON_GOING');
            return;
        }

        oldValueForHandler = oldStatus;
        const isExempt = (s) => s === MDMStatus.ON_GOING || s === MDMStatus.SEND_BACK;
        if (!isExempt(oldStatus) && !isExempt(newStatus)) {
            oldValueForHandler = MDMStatus.ON_GOING;
            const statusRelatedCols = [];
            if (feedbackStatusColIdx > 0) statusRelatedCols.push(ColNames.FEEDBACK_STATUS);
            if (processedDateColIdx > 0) statusRelatedCols.push(ColNames.PROCESSED_DATE);

            if (statusRelatedCols.length) {
                setValuesWithIndexes(sheet, statusRelatedCols, rowIndex, new Array(statusRelatedCols.length).fill(""));
                statusRelatedCols.forEach(colName => { rowData[colName] = ""; });
                mutated = true;
            }
        }
        e.oldValue = oldValueForHandler;
    }

    let activitySheetName = sheetName;
    activitySheetName = getSheetName(rowData[ColNames.REQUEST_TYPE]);
    if (!activitySheetName) {
        Logger.log(`[onChildEdit] WARNING: Unable to resolve Activity Sheet Name for request type: ${rowData[ColNames.REQUEST_TYPE]}`);
        return;
    }

    if (mutated) {
        flushPendingEdits('before handleOnEdit delegation');
    }

    const RequestClass = getRequestClass(activitySheetName);
    const request = new RequestClass(sheet, rowIndex, colIndex, rowData);
    request.handleOnEdit(userEmail, oldValueForHandler);
}

function archive() {
    const archiver = new Archive();
    return archiver.execute();
}

function onChildInterval() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    for (let sheet of sheets) {
        const sheetName = sheet.getName();
        const targetRows = getActivityErrorRow(sheet);
        if (targetRows.length === 0) continue;

        Logger.log(`[onChildInterval] Processing ${sheetName} on rows: ${targetRows.join(', ')}`);

        targetRows.forEach(row => {
            const rowLock = acquireRowLock(sheetName, row, 'onChildInterval', 3, 600);
            if (!rowLock) return;

            try {
                heartbeatRowLock(rowLock);
                const headerRow = sheet.getRange(ACTIVITY_HEADER_ROW_INDEX, 1, 1, sheet.getLastColumn()).getValues()[0];
                const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

                const rowData = {};
                headerRow.forEach((header, index) => {
                    if (header) rowData[header] = rowValues[index];
                });

                let targetSheetName = sheetName;
                const activitySheetName = getSheetName(rowData[ColNames.REQUEST_TYPE]);
                if (activitySheetName) targetSheetName = activitySheetName;
                else return;

                const RequestClass = getRequestClass(targetSheetName);
                const request = new RequestClass(sheet, row, null, rowData);
                const beat = () => { heartbeatRowLock(rowLock); };

                request.handleOnChildInterval(beat);
            } catch (error) {
                Logger.log(`[onChildInterval] Error processing row ${row} in ${sheetName}: ${error.message}`);
            } finally {
                releaseRowLock(rowLock, 'onChildInterval');
            }
        });
    }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('MDM Tools')
        .addSubMenu(ui.createMenu('Business Unit Links')
            .addItem('Retail Unit Alpha', 'openRetailA')
            .addItem('Retail Unit Beta', 'openRetailB')
            .addItem('Industrial Solutions', 'openIndustrial')
            .addItem('Toys & Games', 'openToys')
            .addItem('Food & Beverage', 'openFnB')
            .addSeparator()
            .addItem('Master Activities', 'openMasterProd')
            .addItem('Dashboard', 'openDashboard')
            .addItem('Toolkit', 'openMDMToolkit')
        )
        .addItem('Fix Grant Access', 'fixGrantAccess')
        .addSubMenu(ui.createMenu('Copy Data')
            .addItem('Child (All)', 'copyDataToChildAll')
            .addItem('Child', 'fixSyncToChild')
            .addItem('Master', 'fixSyncToMaster')
        )
        .addItem('Fix On Submit', 'fixOnSubmit')
        .addItem('Sync Attachment Data', 'fixAttachmentSync')
        .addItem('Fix Request Approved', 'fixRequestApproved')
        .addItem("Merge Selected VBS Files", "mergeSelectedVBS")
        .addItem("Set as Special Project", "setDepartmentToSpecialProject")
        .addItem("Change / Add MDM", "changeMDM")
        .addItem("Sort Sheet", "runPrioritySorting")
        .addToUi();
}