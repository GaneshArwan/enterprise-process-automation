/**
 * Represents an email object for sending emails based on activity data.
 * 
 * @param {Sheet} sheet The Google Sheets sheet where the activity data is located.
 * @param {number} rowIndex The row index of the activity within the sheet.
 */
class Email {
    constructor(sheet, rowIndex) {
        this.sheet = sheet;
        this.rowIndex = rowIndex;
        this.activity = new ActivityWithValidation(sheet, rowIndex);
    }

    /**
     * Retrieves the email addresses based on the activity data.
     * 
     * @param {number} [emailRowIndex=null] The row index to use for email retrieval. If not provided, it will be determined dynamically.
     * @returns {Array<string>} An array of email addresses.
     */
    getAllEmails() {
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap();

        let allEmails = [EMAIL_ADDRESS];
        for (let i = 1; i < ATTACHMENT_SYNC_CONTEXTS.length; i++) {
            const ctx = ATTACHMENT_SYNC_CONTEXTS[i];
            const approverEmail = this.getEmailApprover(ctx);

            // Only add approverEmail if it's not undefined or null
            if (approverEmail) {
                allEmails = [...allEmails, ...approverEmail];
            }
        }
        return allEmails;
    }

    getEmailApprover(ctx) {
        const { prop, levelOrder } = ctx;
        const activityValueMap = this.activity.getActivityValueMap();
        const { DEPARTMENT, REQUEST_TYPE } = activityValueMap;
        const companyName = this.activity.getCompanyName();

        const EMAIL_APPROVER = activityValueMap[`EMAIL_${prop}`]
        if (isNotEmpty(EMAIL_APPROVER)) {
            const { valid } = validateEmails(EMAIL_APPROVER);
            if (valid.length) return valid;
        }

        const configApprover = getConfigApproverNew(
            companyName, DEPARTMENT, REQUEST_TYPE, levelOrder
        )

        return configApprover;
    }

    isEmailApprover(ctx) {
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap();
        const emailApprover = this.getEmailApprover(ctx);

        if (emailApprover.length === 0) return false;
        if (emailApprover.includes(EMAIL_ADDRESS)) return false;

        return true;
    }
}

/**
 * Extends the Email class to include email notification functionality.
 * 
 * @param {Sheet} sheet The Google Sheets sheet where the activity data is located.
 * @param {number} rowIndex The row index of the activity within the sheet.
 */
class EmailNotification extends Email {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex); // Call the Email constructor.
    }

    /**
     * Generates the email subject based on the activity data and a given status.
     * 
     * @param {string} subjectStatus The status to include in the email subject.
     * @returns {string} The generated email subject.
     */
    getEmailSubject(subjectStatus) {
        const { REQUEST_TYPE, COMPANY_CODE_NAME, MODIFY_TYPE, REQUEST_NUMBER } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const requestType = MODIFY_TYPE ? `${REQUEST_TYPE} (${MODIFY_TYPE})` : REQUEST_TYPE
        return (
            `${subjectStatus} ${requestType} Submission ${COMPANY_CODE_NAME} ${REQUEST_NUMBER}`
        ); // Construct the email subject.
    }

    /**
     * Generates the email body based on the activity data and additional content.
     * 
     * @param {string} additionalContent Additional content to include in the email body.
     * @returns {string} The generated email body.
     */
    getEmailBody(recipient, additionalContent) {
        const {
            REQUEST_TYPE, REQUEST_NUMBER, COMPANY_CODE_NAME,
            EMAIL_ADDRESS, DEPARTMENT, DOCUMENT_NUMBER, MODIFY_TYPE,
            TRANSACTION_SECTION
        } = this.activity.getActivityValueMap(); // Retrieve activity data.

        const requestDetails = [
            `<strong>Request No:</strong> ${REQUEST_NUMBER}`,
            `<strong>Business Unit:</strong> ${COMPANY_CODE_NAME}`,
            DEPARTMENT ? `<strong>Department:</strong> ${DEPARTMENT}` : null,
            `<strong>Request Type:</strong> ${REQUEST_TYPE}`,
            `<strong>Submission by:</strong> ${EMAIL_ADDRESS}`,
            DOCUMENT_NUMBER ? `<strong>Document Number:</strong> ${DOCUMENT_NUMBER}` : null,
            MODIFY_TYPE ? `<strong>Modify Type:</strong> ${MODIFY_TYPE}` : null,
            TRANSACTION_SECTION ? `<strong>Transaction Section:</strong> ${TRANSACTION_SECTION}` : null,
        ]
            .filter(Boolean)
            .map(detail => createStyledParagraph(detail))
            .join('');

        const body = `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                ${createStyledParagraph(`Dear ${recipient},`)}
                ${createStyledParagraph(`This is a confirmation message about ${REQUEST_TYPE} submission.`)}
                ${requestDetails}
                ${additionalContent}
                ${createStyledParagraph(`<span style="color: #666; font-style: italic;">
                    Please don't reply to this email</span>`)}
            </div>
        `;

        return body;
    }

    /**
     * Sends an email based on the provided parameters.
     * 
     * @param {Array<string>} emails The email addresses to send to.
     * @param {string} recipient The recipient's name.
     * @param {string} subjectStatus The status to include in the email subject.
     * @param {string} additionalContent Additional content to include in the email body.
     * @param {Object} [options={}] Optional parameters for the email.
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmail(emails, recipient, subjectStatus, additionalContent, options = {}) {
        console.log("EMAILS: ", emails)

        try {
            const body = this.getEmailBody(recipient, additionalContent.split('\n')
                .filter(line => line.trim())
                .map(line => createStyledParagraph(line))
                .join(''));

            const emailConfig = {
                to: emails.join(', '),
                subject: this.getEmailSubject(subjectStatus),
                body: body.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim(),
                htmlBody: body,
                ...options
            };

            try {
                GmailApp.sendEmail(emailConfig.to, emailConfig.subject, emailConfig.body, emailConfig);
            } catch {
                MailApp.sendEmail(emailConfig);
            }

            Logger.log("Email successfully sent to: " + emailConfig.to)
            return true;
        } catch (error) {
            console.error("Failed to send email:", error);
            return false;
        }
    }
}

/**
 * Extends the EmailNotification class to include email handling functionality.
 * 
 * @param {Sheet} sheet The Google Sheets sheet where the activity data is located.
 * @param {number} rowIndex The row index of the activity within the sheet.
 */
class EmailHandler extends EmailNotification {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex); // Call the EmailNotification constructor.
    }

    /**
     * Sends an email for a new request submission.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailNewRequest() {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody = `A new request has been submitted. Please fill out the attachment for submission\n` +
            `${ATTACHMENT}`;
        const additionalSubject = 'New Request Submission';
        const recipient = 'Requester';

        console.log("EMAIL ADDRESS: ", EMAIL_ADDRESS)
        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailAskApproval(ctx) {
        const { ATTACHMENT } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody = `A new request needs your approval. Please review and approve it using the link below:\n` +
            `${ATTACHMENT}`;
        const additionalSubject = 'Request for Approval';
        const recipient = 'Approver';

        const emailAddress = this.getEmailApprover(ctx) // Get the email address of the approver.

        return this.sendEmail(
            emailAddress,
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email after a request is approved.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailApproved() {
        const { REQUEST_TYPE, ATTACHMENT, RESPON_REQUESTER, RESPON_APPROVER } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody = `Please continue this request in spreadsheet: \n${ATTACHMENT}`;
        const additionalSubject = `Request ${RESPON_APPROVER ? RESPON_APPROVER : RESPON_REQUESTER}`;
        const recipient = 'MDM Team';
        const email = [EMAIL_MDM_GROUP];

        let ccEmail;
        switch (REQUEST_TYPE) {
            case RequestTypes.PROMOTION_CREATE:
                const promoCode = this.activity.getPromoCode();
                if (promoCode !== 'Z005') return
                ccEmail = EMAIL_CC_SECOND;
                break;
            default:
                return;
        }

        return this.sendEmail(
            email,
            recipient,
            additionalSubject,
            additionalBody,
            { cc: ccEmail }
        );
    }

    /**
     * Sends an email when a request expires.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailExpired() {
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `Your request has expired due to inactivity. Please submit a new request if necessary`;
        const additionalSubject = "Request Expired";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email when a request is rejected.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailRejected(ctx) {
        const { NAME_APPROVER, ATTACHMENT, TIMESTAMP_APPROVAL, EMAIL_ADDRESS } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `The request has been rejected by ${NAME_APPROVER} on ${TIMESTAMP_APPROVAL}.\n` +
            `For more information, please reach out to the approver.\n` +
            `Link attachment: ${ATTACHMENT}`
        const additionalSubject = 'Request Rejected by Approver';
        const recipient = 'Requester';
        const emailAddress = this.getEmailApprover(ctx)

        return this.sendEmail(
            [EMAIL_ADDRESS,
                ...emailAddress
            ],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailInvalid(ctx) {
        const { EMAIL_ADDRESS, ATTACHMENT } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `Request Sent Back due to invalid status/name.\n` +
            `Please revise your input using the link below:\n` +
            `${ATTACHMENT}`

        const additionalSubject = "Request Sent Back Invalid Status/Name";
        let recipient = 'Requester'
        let targetEmail = EMAIL_ADDRESS

        if (ctx.levelOrder > 0) {
            recipient = "Approver"
            targetEmail = this.getEmailApprover(ctx);
            if (targetEmail.length === 0) return;
        }

        return this.sendEmail(
            targetEmail,
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email for non-valid requester status.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailInvalidRequester() {
        const { EMAIL_ADDRESS, ATTACHMENT } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `The requester status/name is invalid.\n` +
            `Please revise your input using the link below:\n` +
            `${ATTACHMENT}`

        const additionalSubject = "Request Sent Back for Review";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailInvalidApprover(ctx) {
        const { ATTACHMENT } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `The approver status/name is invalid.\n` +
            `Please revise your input using the link below:\n` +
            `${ATTACHMENT}`
        const additionalSubject = 'Approval Failed';
        const recipient = 'Approver';

        return this.sendEmail(
            this.getEmailApprover(ctx),
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email for non-valid approver status.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailNonValidApprover() {
        const { ATTACHMENT } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `The approver status/name is invalid.\n` +
            `Please revise your input using the link below:\n` +
            `${ATTACHMENT}`
        const additionalSubject = 'Approval Failed';
        const recipient = 'Approver';

        return this.sendEmail(
            this.getEmailApprover(),
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailNonvalidRequester() {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `Requester status/name is invalid.\n` +
            `Please revise your input using the link below:\n` +
            `${ATTACHMENT}`
        const additionalSubject = 'Requester Failed';

        return this.sendEmail(
            EMAIL_ADDRESS,
            'Requester',
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email when there is an empty task.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailEmptyTask() {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody =
            `No task is currently assigned in the sheet.\n` +
            `Please complete the task at the link below before submission:\n` +
            `${ATTACHMENT}`

        const additionalSubject = "Request Sent Back for Review";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email when there is an empty mandatory field.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailEmptyMandatory() {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody =
            `Mandatory fields are still incomplete. \n` +
            `Please fill out all required columns using the link below:\n` +
            `${ATTACHMENT}`

        const additionalSubject = "Request Sent Back Empty Mandatory";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailSendBackBySystem(reason) {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody =
            `\n-------\n` +
            `Request are sent back automatically by System. \n` +
            `${reason ? `reason: ${reason}` : ''}\n` +
            `Please check 'REMARKS(MDM)' column within attachment for Issues detail. \n` +
            `After revision, re-trigger Complete Status within the attachment below:\n` +
            `${ATTACHMENT}`

        const additionalSubject = "Request Sent Back By System";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailSendBackByApprover(name, reason) {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody = `
        -------  
        Request are sent back by Approver.  
        ${name ? `<strong>Name:</strong> ${name}<br>` : ''}  
        ${reason ? `<strong>Reason:</strong> ${reason}<br>` : 'Contact Approver for further detail<br>'}  
        After revision, re-trigger Complete Status within the attachment below:<br>  
        ${ATTACHMENT}
        `;

        const additionalSubject = "Request Sent Back By Approver";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    sendEmailSendBackByMDM(remark) {
        const { ATTACHMENT, EMAIL_ADDRESS } = this.activity.getActivityValueMap(true,true); // Retrieve activity data.
        const additionalBody = `
        -------  
        Request is sent back by MDM.  
        Remark: ${remark ? remark : 'Check REMARKS(MDM) column for further detail'}

        After revision, re-trigger Complete Status (C17) within the attachment below:<br>  
        ${ATTACHMENT}
        `;

        const additionalSubject = "Request Sent Back By MDM";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );
    }

    /**
     * Sends an email after a request is processed.
     * 
     * @returns {boolean} True if the email is sent successfully, false otherwise.
     */
    sendEmailProcessed() {
        const {
            REQUEST_NUMBER, PROCESS_STATUS, NAME_APPROVER,
            PROCESSED_BY, PROCESSED_DATE, EMAIL_ADDRESS,
            ATTACHMENT, REMARK,
        } = this.activity.getActivityValueMap(); // Retrieve activity data.

        const FEEDBACK_URL =
            // (a) the exact “viewform” endpoint:
            "https://docs.google.com/forms/d/e/"
            + "1FAIpQLSe3eo8c_s0qXKLoVrrh8HiMBJgbKyqWV4hIfsUVuui8C0UG1Q"
            + "/viewform"
            // (b) tell Google this is a “pre-filled” URL:
            + "?usp=pp_url"
            // (c) append our entry.<ID>=VALUE, URL-encoded
            + "&entry.136009593="
            + encodeURIComponent(REQUEST_NUMBER);

        const processDetails = [
            `<strong>Status:</strong> ${PROCESS_STATUS}`,
            // Only show “Approver” if NAME_APPROVER is truthy
            NAME_APPROVER ? `<strong>Approver:</strong> ${NAME_APPROVER}` : '',
            `<strong>Processed by:</strong> ${PROCESSED_BY? PROCESSED_BY : this.sheet.getName()}`,
            `<strong>Processed Date:</strong> ${PROCESSED_DATE}`,

            // If there’s a REMARK, insert two separate lines
            ...(
                REMARK
                    ? [
                        `<strong>MDM Notes:</strong> ${REMARK}`,
                        `<strong>Please check the attachment below for other remark details:</strong>`
                    ]
                    : []
            ),

            ATTACHMENT,
            '---------------------------------',

            // Now include the feedback prompt + a link that has entry.136009593=REQUEST_NUMBER
            `<strong>For our improvement, please provide your feedback regarding our operations on the link below:</strong>`,
            `<a href="${FEEDBACK_URL}">${FEEDBACK_URL}</a>`
        ]
            .filter(Boolean)
            .map(detail => createStyledParagraph(detail))
            .join('');

        const feedbackValues = new Attachment(this.sheet, this.rowIndex).getFeedbackValues();

        let additionalContent = processDetails;

        if (feedbackValues && PROCESS_STATUS !== ApproverStatus.REJECTED) {
            const tableHTML = createTableHTML(feedbackValues);
            additionalContent += `
                <div style="margin-top: 20px;">
                    <p style="margin: 5px 0;"><strong>Below is the table of MDM feedback:</strong></p>
                    ${tableHTML}
                </div>
            `;
        }

        return this.sendEmail(
            [EMAIL_ADDRESS],
            'Requester',
            `Request ${PROCESS_STATUS}`,
            additionalContent,
        );
    }

    sendEmailInvalidRequester() {
        const { EMAIL_ADDRESS } = this.activity.getActivityValueMap(); // Retrieve activity data.
        const additionalBody =
            `You are not allowed to perform this Request. \n` +
            `Please contact your MDM if you should be allowed.\n`

        const additionalSubject = "Invalid Requester";
        const recipient = 'Requester';

        return this.sendEmail(
            [EMAIL_ADDRESS],
            recipient,
            additionalSubject,
            additionalBody
        );

    }
}

class EmailHandlerVendor extends EmailHandler {
    constructor(sheet, rowIndex) {
        super(sheet, rowIndex);
    }

    getEmailApprover(ctx) {
        const { prop, levelOrder } = ctx;
        const activityValueMap = this.activity.getActivityValueMap();
        const { MODIFY_TYPE, BANK_TYPE, TRANSACTION_SECTION, DEPARTMENT, REQUEST_TYPE } = activityValueMap;
        const companyCode = this.activity.getCompanyName();

        let lookupRequestType = REQUEST_TYPE;
        if (isNotEmpty(MODIFY_TYPE))
            lookupRequestType = `${REQUEST_TYPE} (${MODIFY_TYPE})`;
        if (isNotEmpty(BANK_TYPE))
            lookupRequestType = `${REQUEST_TYPE} (${BANK_TYPE})`;


        if (isNotEmpty(TRANSACTION_SECTION)) {
            const lookupDepartment = `${TRANSACTION_SECTION} (AP)`;
            const approver = getConfigApproverNew(
                companyCode, lookupDepartment, lookupRequestType, ctx.levelOrder, false)

            if (approver.length > 0) return approver
        }

        let approver = getConfigApproverNew(
            companyCode, DEPARTMENT, lookupRequestType, ctx.levelOrder, true)

        const EMAIL_APPROVER = activityValueMap[`EMAIL_${prop}`]
        if (isNotEmpty(EMAIL_APPROVER)) {
            const { valid } = validateEmails(EMAIL_APPROVER);
            if (valid.length) { 
                approver = [...approver, ...valid];
            }
        }

        return approver

    }
}