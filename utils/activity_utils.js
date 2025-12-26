function getActivityRowToCrawl(sheet) {
    const ACT = ACTIVITY_HEADER_ROW_INDEX;

    const approverCols = ATTACHMENT_SYNC_CONTEXTS
        .slice(1)
        .map(item => ColNames[`RESPON_${item.prop}`]);

    const [reqNos, atts, reqs, ...apprs] = getValuesByColumns(
        sheet,
        [
            ColNames.REQUEST_NUMBER,
            ColNames.ATTACHMENT,
            ColNames.RESPON_REQUESTER,
            ...approverCols
        ],
        ACT
    );

    const REJ = String(ApproverStatus.REJECTED).trim().toLowerCase();

    const result = reqs
        .map((req, i) => {
            const row = i + ACT;

            // skip expired/invalid or missing request-number / attachment
            if (
                req === RequesterStatus.EXPIRED ||
                req === RequesterStatus.INVALID ||
                !reqNos[i] ||
                !atts[i]
            ) {
                return null;
            }

            // ▶︎ if the requester cell is empty, always crawl this row
            if (!req || req == RequesterStatus.NEED_REVIEW) {
                return { row, requestNumber: reqNos[i] };
            }

            // otherwise, scan approver columns for a “pending” (blank) before
            // hitting undefined (end-of-data) or a “rejected” marker
            const vals = apprs.map(col => col[i]);

            for (const v of vals) {
                if (typeof v === 'undefined') {
                    break;
                }
                const txt = String(v).trim().toLowerCase();
                if (txt === REJ) {
                    break;
                }
                if (txt === '') {
                    return { row, requestNumber: reqNos[i] };
                }
            }

            return null;
        })
        .filter(r => r && r.row > ACT)
        .sort((a, b) => b.row - a.row);

    return result;
}


function getNewSubmissionStatusError(sheet) {
    const [requestNumber, attachment, newSubmissionStatus, responRequester] = getValuesByColumns(
        sheet,
        [
            ColNames.REQUEST_NUMBER,
            ColNames.ATTACHMENT,
            ColNames.NEW_SUBMISSION_STATUS,
            ColNames.RESPON_REQUESTER
        ],
        ACTIVITY_HEADER_ROW_INDEX
    )

    return newSubmissionStatus.reduce((emptyRows, submissionStatus, i) => {
        if (requestNumber[i] && attachment[i] && !submissionStatus && !responRequester[i]) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX)
        }

        return emptyRows
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a)
}

function getAskApprovalStatusError(sheet) {
    const [responRequester, nameApprover, askApprovalStatus] = getValuesByColumns(
        sheet,
        [
            ColNames.RESPON_REQUESTER,
            ColNames.NAME_APPROVER,
            ColNames.ASK_APPROVAL_STATUS
        ],
        ACTIVITY_HEADER_ROW_INDEX
    )

    return askApprovalStatus.reduce((emptyRows, approvalStatus, i) => {
        if (
            !approvalStatus && !nameApprover[i] &&
            (responRequester[i] && ![RequesterStatus.EXPIRED, RequesterStatus.INVALID, RequesterStatus.NEED_REVIEW].includes(responRequester[i]))
        ) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX)
        }

        return emptyRows
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a)
}

function getAskApprovalFinalStatusError(sheet) {
    const [responApprover, nameApproverFinal, askApprovalFinalStatus] = getValuesByColumns(
        sheet,
        [
            ColNames.RESPON_APPROVER,
            ColNames.NAME_APPROVER_FINAL,
            ColNames.ASK_APPROVAL_FINAL_STATUS
        ],
        ACTIVITY_HEADER_ROW_INDEX
    )

    return askApprovalFinalStatus.reduce((emptyRows, approvalFinalStatus, i) => {
        if (
            !approvalFinalStatus && !nameApproverFinal[i] && responApprover[i] !== ApproverStatus.REJECTED &&
            (responApprover[i] && ![RequesterStatus.EXPIRED, RequesterStatus.INVALID, RequesterStatus.NEED_REVIEW].includes(responApprover[i]))
        ) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX)
        }

        return emptyRows
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a)
}

function getActivityErrorRow(sheet) {
    const [
        takenDate, processStatus,
        processedDate, feedbackStatus,
        estimatedTimeFinished, estimatedTime, attachment,
        noArToSap,mdmApprovalDate
    ] = getValuesByColumns(
        sheet,
        [
            ColNames.TAKEN_DATE,
            ColNames.PROCESS_STATUS,
            ColNames.PROCESSED_DATE,
            ColNames.FEEDBACK_STATUS,
            ColNames.ESTIMATED_TIME_FINISHED,
            ColNames.ESTIMATED_TIME,
            ColNames.ATTACHMENT,
            ColNames.NO_AR_TO_SAP,
            ColNames.MDM_APPROVAL_DATE,
        ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    const errorRows = processStatus.reduce((emptyRows, status, i) => {
        const rowIndex = i + ACTIVITY_HEADER_ROW_INDEX;
        
        // Aturan error untuk feedback standar & master site.
        let isErrFeedback = false;
        if (!noArToSap[i]) {
            isErrFeedback = status && status !== MDMStatus.ON_GOING && processedDate[i] && !feedbackStatus[i];
        }else{
            isErrFeedback = mdmApprovalDate[i] && !feedbackStatus[i];
        }

        // Aturan error untuk estimated time.
        const isErrEstimatedTime = takenDate[i] && estimatedTime[i] && !estimatedTimeFinished[i];
        
        // Aturan error untuk "Send Back".
        let isErrSendBack = status === MDMStatus.SEND_BACK && !feedbackStatus[i];
        if (attachment[i] === "NO ATTACHMENT") {
            isErrSendBack = false;
        }

        // Gabungkan semua kondisi error.
        if (isErrFeedback || isErrEstimatedTime || isErrSendBack) {
            emptyRows.push(rowIndex);
        }

        return emptyRows;
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a);
        
    Logger.log(`[getActivityErrorRow] Sheet ${sheet.getName()}: Found ${errorRows.length} error rows: [${errorRows.join(', ')}]`);
    return errorRows;
}

function getOnSubmitErrorRow(sheet) {
    const [requestNum, timestamp, attachment] = getValuesByColumns(
        sheet,
        [
            ColNames.REQUEST_NUMBER,
            ColNames.TIMESTAMP,
            ColNames.ATTACHMENT
        ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    return requestNum.reduce((emptyRows, req, i) => {

        //Timed Out is 5 minutes after first requested
        const isTimedOut = timestamp[i] && getMinuteDiff(timestamp[i]) > 10;

        if ((!req || !attachment[i]) && isTimedOut) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX);
        }
        return emptyRows
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a);
}

function getSystemSentBackErrorRow(sheet) {
    const [responRequester, sentBackCount, sentBackEmailStatus] = getValuesByColumns(
        sheet,
        [
            ColNames.RESPON_REQUESTER,
            ColNames.SYSTEM_SENT_BACK_COUNT,
            ColNames.SYSTEM_SENT_BACK_EMAIL_STATUS
        ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    return sentBackCount.reduce((errorRows, count, i) => {
        if (responRequester[i] === RequesterStatus.NEED_REVIEW) {
            if (!count) return errorRows;
            const statusValue = String(sentBackEmailStatus[i] || ""); // Ensure it's a string
            const cleanedStatus = statusValue
                .split(SYSTEM_SENT_BACK_SEPARATOR)
                .filter(status => status.trim() !== '');

            if (parseInt(count, 10) !== cleanedStatus.length) {
                errorRows.push(i + ACTIVITY_HEADER_ROW_INDEX); // Adjust for row index
            }
        }
        return errorRows;
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a);
}

function getEmptyTotalTaskRow(sheet) {
    const [totalTask, requester, approver, approverFinal] = getValuesByColumns(
        sheet,
        [
            ColNames.TOTAL_TASK,
            ColNames.RESPON_REQUESTER,
            ColNames.RESPON_APPROVER,
            ColNames.RESPON_APPROVER_FINAL
        ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    return totalTask.reduce((emptyRows, row, i) => {

        const validApprover = approver?.length > 0
            ? approverFinal?.length > 0
                ? approverFinal[i]
                : approver[i]
            : requester[i] && requester[i] !== RequesterStatus.EXPIRED;


        if (!row && validApprover) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX);
        }

        return emptyRows;
    }, [])
        .filter(row => row > 2)
        .sort((a, b) => b - a);
}

function getSheetName(requestType) {
    return Object.entries(RequestTypeActivityMap).find(
        ([, requestTypes]) => requestTypes.includes(requestType)
    )?.[0] || null;
}



function getRowsWithTimestampEntryButNoTakenDate(sheet) {
    const [timestampEntry, takenDate] = getValuesByColumns(
        sheet,
        [
            ColNames.TIMESTAMP_ENTRY,
            ColNames.TAKEN_DATE
        ],
        ACTIVITY_HEADER_ROW_INDEX
    );

    return timestampEntry.reduce((emptyRows, timestamp, i) => {
        // Row has timestamp entry but no taken date
        if (timestamp && !takenDate[i]) {
            emptyRows.push(i + ACTIVITY_HEADER_ROW_INDEX);
        }
        return emptyRows;
    }, [])
        .filter(row => row > ACTIVITY_HEADER_ROW_INDEX)
        .sort((a, b) => b - a);
}
