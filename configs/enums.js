const ScriptTypes = Object.freeze({
    SCRIPT_RETAIL_NEW: 'ScriptRetailNew',
    SCRIPT_TOYS_CONS: 'ScriptToysCons',
    SCRIPT_GENERIC_CONS: 'ScriptGenericCons',
    NO_SITE: 'NoSite',
    UPDATE_PIR_AND_CURRENCY: 'UpdatePirAndCurrency',
    UPDATE_PIR_NEW: 'UpdatePIRNew',
    CURRENCY_NEW: 'CurrencyNew',
    SCRIPT_HOME_ESSENTIALS: 'ScriptHomeEssentials',
    SCRIPT_MANUF: 'ScriptManuf',
    SCRIPT_SAP_A: 'ScriptSapA',
    SCRIPT_SAP_B: 'ScriptSapB',
    SCRIPT_MODIFY_PRICE: 'ScriptModifyPrice',
    CREATE_PIR_SITE: 'CreatePIRSite',
    CREATE_PIR_SITE_UNIT: 'CreatePIRSiteUnit'
})

const SystemActor = Object.freeze({
    SYSTEM: 'SYSTEM',
    APPROVER: 'APPROVER',
    MDM: 'MDM'
})

const ActivityLog = Object.freeze({
    SEND_BACK: 'Send Back'
})

const ColNames = Object.freeze({
    REQUEST_NUMBER: 'Request Number',
    TIMESTAMP: 'Timestamp',
    EMAIL_ADDRESS: 'Email Address',
    EMAIL_APPROVER: 'Email Approver',
    REQUESTER_NAME: 'Requester Name',
    ATTACHMENT: 'Attachment',
    REQUEST_TYPE: 'Request Type',
    DEPARTMENT: 'Department',
    RESPON_REQUESTER: 'Respon Requester',
    TIMESTAMP_REQUESTER: 'Timestamp Requester',
    NAME_REQUESTER: 'Name Requester',
    RESPON_APPROVER: 'Respon Approver',
    RESPON_APPROVER_II: 'Respon Approver II',
    RESPON_APPROVER_III: 'Respon Approver III',
    TIMESTAMP_APPROVER: 'Timestamp Approver',
    TIMESTAMP_APPROVER_II: 'Timestamp Approver II',
    TIMESTAMP_APPROVER_III: 'Timestamp Approver III',

    NAME_APPROVER: 'Name Approver',
    NAME_APPROVER_II: 'Name Approver II',
    NAME_APPROVER_III: 'Name Approver III',
    TOTAL_TASK: 'Total Task',
    TOTAL_PROMO: 'Total Promo',
    FEEDBACK_STATUS: 'Feedback Status',
    PROCESSED_BY: 'Processed By',
    PROCESS_STATUS: 'Process Status',
    TAKEN_DATE: 'Taken Date',
    PROCESSED_DATE: 'Processed Date',
    ADDITIONAL_ATTACHMENT: 'Additional Attachment',
    DOCUMENT_NUMBER: 'Document Number',
    REQUEST_COUNT: 'COUNTER',
    MDM_APPROVAL_DATE: 'MDM Approval Date',
    MDM_PIC_AR: 'MDM Approved By',
    NO_AR_TO_SAP: 'No AR to SAP',
    NEW_SUBMISSION_STATUS: 'New Submission Status',
    ASK_APPROVER_STATUS: 'Ask Approver Status',
    REMARK: 'Remark',
    ACCESS_SEQUENCE: 'Access Sequence',
    ASK_APPROVER_III_STATUS: 'Ask Approver III Status',
    ASK_APPROVER_II_STATUS: 'Ask Approver II Status',
    CHILD_REQUEST_NUMBER: 'Child Request Number',
    SYSTEM_SENT_BACK_COUNT: 'System Sent Back Count',
    SYSTEM_SENT_BACK_TIME: 'System Sent Back Time',
    SYSTEM_SENT_BACK_ACTOR: 'System Sent Back Actor',
    SYSTEM_SENT_BACK_EMAIL_STATUS: 'System Sent Back Email Status',
    SCRIPT_TYPE: 'Script Type',
    SCRIPT_FILE: 'Script File',
    BANK_TYPE: 'Bank Type',

    TIMESTAMP_ENTRY: 'Timestamp Entry',
    VALID_FROM: 'Valid From',
    BASELINE: 'Baseline',
    ESTIMATED_TIME: 'Estimated Time',
    ESTIMATED_TIME_FINISHED: 'Estimated Time Finished',
})

const ActivitySheetNames = Object.freeze({
    IMAGE: 'IMAGE',
    MERCHANDISE: 'MERCHANDISE',
    PROMOTION: 'PROMOTION',
    STATUS_LISTING: 'STATUS/LISTING',
    MASTER_DATA: 'MASTER DATA',
    EXTEND_PIR: 'EXTEND PIR',
    SOURCE_LIST: 'SOURCE LIST',
    HIERARCHY: 'HIERARCHY',
    BASIC_DATA: 'BASIC DATA',
    NON_M: 'NON M',
    BOM: 'BOM',
    MASTER_FINANCE: 'MASTER FINANCE',
    MASTER_SITE: 'MASTER SITE',
    PRICING: 'PRICING',
    PROFIT_CENTER: 'PROFIT CENTER',
    CUSTOMER: 'CUSTOMER',
    VENDOR: 'VENDOR'
})

// Sanitized Agent Names
const MDMSheetNames = Object.freeze({
    AGENT_01 : 'AGENT_01',
    AGENT_02 : 'AGENT_02',
    AGENT_03 : 'AGENT_03',
    AGENT_04 : 'AGENT_04',
    AGENT_05 : 'AGENT_05',
    AGENT_06 : 'AGENT_06',
    AGENT_07 : 'AGENT_07',
    AGENT_08 : 'AGENT_08',
    AGENT_09 : 'AGENT_09',
    AGENT_10 : 'AGENT_10',
    AGENT_11 : 'AGENT_11',
    AGENT_12 : 'AGENT_12'
})

const RequestTypes = Object.freeze({
    BLOCK: 'Block',
    BLOCK_TEMPORARY: 'Block Temporary',
    UNBCLOCK: 'Unblock',
    UNBLOCK_TEMPORARY: 'Unblock Temporary',
    DELISTING: 'Delisting',
    INACTIVE: 'Inactive',
    LISTING: 'Listing',
    LISTING_TEMPORARY: 'Listing Temporary',
    SLOC: 'Sloc',
    PROMOTION_CREATE: 'Promotion Create',
    PROMOTION_MODIDY_CHANGE: 'Promotion Modify/Change',
    MERCHANDISE_CREATE_NO_IMAGE: 'Create Article Merchandise Without Image / Picture',
    MERCHANDISE_CREATE_IMAGE: 'Create Article Merchandise With Image / Picture',
    MASTER_DATA_CREATE: 'Master Data Create',
    MASTER_DATA_MODIFY_CHANGE: 'Master Data Modify/Change',
    IMAGE_MODIFY_CHANGE: 'Image Modify/Change',
    PIR_CREATE: 'PIR Create',
    EXTEND_PIR: 'Extend and PIR',
    PIR_MODIFY_CHANGE: 'PIR Modify/Change',
    HIERARCHY_RECLASS: 'Article Hierarchy Reclass',
    HIERARCHY_CREATE: 'Article Hierarchy Create',
    HIERARCHY_MODIFY_CHANGE: 'Article Hierarchy Modify/Change',
    BOM_CREATE: 'BOM Create',
    BOM_MODIFY_CHANGE: 'BOM Modify/Change',
    NON_M_CREATE: 'Non M Create',
    BASIC_DATA_MODIFY_CHANGE: 'Basic Data Modify/Change',
    SOURCE_LIST: 'Source List',

    //Master Finance
    GL_CREATE: 'GL Create',
    GL_MODIFY_CHANGE: 'GL Modify/Change',
    GL_BLOCK: 'GL Block',
    GL_UNBLOCK: 'GL Unblock',
    GL_EXTEND: 'GL Extend',
    COST_CENTER_CREATE: 'Cost Center Create',
    COST_CENTER_MODIFY_CHANGE: 'Cost Center Modify/Change',
    COST_CENTER_BLOCK: 'Cost Center Block',
    COST_CENTER_UNBLOCK: 'Cost Center Unblock/Undelete',
    COST_CENTER_UNBLOCK_TEMPORARY: 'Cost Center Unblock Temporary',
    COST_CENTER_UNBLOCK_BLOCK: 'Cost Center Unblock Temporary (Block)',

    //Master Site
    SITE_CREATE: 'Site Master Create',
    SITE_MODIFY_CHANGE: 'Site Master Modify/Change (Name/Search Term 2)',
    SITE_MODIFY_CHANGE_OTHERS: 'Site Master Modify/Change (Others)',
    SITE_BLOCK: 'Site Master Block',
    SITE_UNBLOCK: 'Site Master Unblock',

    //Profit Center
    PROFIT_CENTER_BLOCK: 'Profit Center Block',
    PROFIT_CENTER_UNBLOCK: 'Profit Center Unblock',

    SALES_PERSON_CREATE: 'Sales Person Create',
    SALES_PERSON_MODIFY_CHANGE: 'Sales Person Modify/Change',
    SALES_PERSON_BLOCK: 'Sales Person Block',
    CUSTOMER_CREATE_BADAN_USAHA: 'Customer Create (Badan Usaha)',
    CUSTOMER_CREATE_PERORANGAN: 'Customer Create (Perorangan)',
    CUSTOMER_MODIFY_CHANGE: 'Customer Modify/Change',
    CUSTOMER_EXTEND: 'Customer Extend',
    CREDIT_LIMIT_CREATE_UPDATE: 'Credit Limit Create/Update',
    TOP_UPDATE: 'TOP Update',
    TOP_CREDIT_LIMIT_UPDATE: 'TOP & Credit Limit Update',
    CUSTOMER_BLOCK: 'Customer Block',
    CUSTOMER_UNBLOCK: 'Customer Unblock',
    CUSTOMER_CREATE: 'Customer Create',

    //Pricing
    PRICING_CREATE: 'Pricing Create',
    PRICING_MODIFY_CHANGE: 'Pricing Modify/Change',

    BANK_KEY_MASTER_MODIFY_CHANGE: 'Bank Key Master Modify/Change',
    BANK_KEY_MASTER_CREATE: 'Bank Key Master Create',
    BANK_KEY_MASTER_DELETE: 'Bank Key Master Delete',
    VENDOR_MODIFY_CHANGE: 'Vendor Modify/Change',
    ACCOUNT_VENDOR_MODIFY_CHANGE: 'Account Vendor Modify/Change',
    ACCOUNT_VENDOR_ADD : 'Account Vendor Add',
    ACCOUNT_VENDOR_DELETE : 'Account Vendor Delete',
    RETURN_TO_VENDOR: 'Return To Vendor',
    VENDOR_EXTEND: 'Vendor Extend',
    VENDOR_BLOCK: 'Vendor Block'
})

const SheetNames = Object.freeze({
    REQUEST: 'TOTAL REQUEST',
    MDM_FEEDBACK: 'MDM FEEDBACK',
    IMAGE: 'IMAGES',
    LINK_BLOCK: 'LINK BLOCK',
})

let AttachmentValues = Object.freeze({
    COMPANY_CELL: 'F10',
    REQUESTER_STATUS_CELL: 'C17',
    REQUESTER_NAME_CELL: 'C18',
    REQUESTER_NOTES_CELL: 'C19',
    APPROVER_STATUS_CELL: 'D17',
    APPROVER_NAME_CELL: 'D18',
    APPROVER_NOTES_CELL: 'D19',
    APPROVER_II_STATUS_CELL: 'E17',
    APPROVER_II_NAME_CELL: 'E18',
    APPROVER_II_NOTES_CELL: 'E19',
    APPROVER_III_STATUS_CELL: 'F17',
    APPROVER_III_NAME_CELL: 'F18',
    APPROVER_III_NOTES_CELL: 'F19',
    IMAGE_CELL: 'B22',
    TASK_START_ROW: 25,
    UID_COL_INDEX: 3,
    MANDATORY_COLOR: '#ffff00',
})

const RequesterStatus = Object.freeze({
    COMPLETED: "Completed",
    EXPIRED: "Expired",
    INVALID: "Invalid",
    NEED_REVIEW: 'Need Review'
})

const ApproverStatus = Object.freeze({
    APPROVED: "Approved",
    REJECTED: "Rejected",
    PARTIALLY_REJECTED: "Partially Rejected",
    SEND_BACK: "Send Back"
})

const MasterConfiguration = Object.freeze({
    DRIVE_SUFFIX: "_DRIVE",
    IMAGE_DRIVE_SUFFIX: "_IMAGE",
    ADDITIONAL_SUFFIX: "_ADDITIONAL",
    CHILD_SPREADSHEET_KEY: "SPREADSHEET",
})

const MDMStatus = Object.freeze({
    COMPLETED: 'Completed',
    PARTIALLY_REJECTED: 'Partially Rejected',
    REJECTED: 'Rejected',
    SEND_BACK: 'Send Back',
    ON_GOING: 'On Going',
})

const RequestTypeActivityMap = {
    [ActivitySheetNames.MERCHANDISE]: [
        RequestTypes.MERCHANDISE_CREATE_IMAGE,
        RequestTypes.MERCHANDISE_CREATE_NO_IMAGE
    ],

    [ActivitySheetNames.IMAGE]: [
        RequestTypes.IMAGE_MODIFY_CHANGE
    ],

    [ActivitySheetNames.BASIC_DATA]: [
        RequestTypes.BASIC_DATA_MODIFY_CHANGE
    ],

    [ActivitySheetNames.STATUS_LISTING]: [
        RequestTypes.UNBLOCK_TEMPORARY,
        RequestTypes.BLOCK, 
        RequestTypes.INACTIVE, 
        RequestTypes.DELISTING,
        RequestTypes.UNBCLOCK,
        RequestTypes.BLOCK_TEMPORARY,
        RequestTypes.LISTING,
        RequestTypes.LISTING_TEMPORARY,
        RequestTypes.SLOC
    ],

    [ActivitySheetNames.PROMOTION]: [
        RequestTypes.PROMOTION_CREATE,
        RequestTypes.PROMOTION_MODIDY_CHANGE
    ],

    [ActivitySheetNames.MASTER_DATA]: [
        RequestTypes.MASTER_DATA_CREATE,
        RequestTypes.MASTER_DATA_MODIFY_CHANGE
    ],

    [ActivitySheetNames.EXTEND_PIR]: [
        RequestTypes.PIR_CREATE,
        RequestTypes.EXTEND_PIR,
        RequestTypes.PIR_MODIFY_CHANGE
    ],

    [ActivitySheetNames.HIERARCHY]: [
        RequestTypes.HIERARCHY_CREATE,
        RequestTypes.HIERARCHY_MODIFY_CHANGE,
        RequestTypes.HIERARCHY_RECLASS
    ],

    [ActivitySheetNames.BOM]: [
        RequestTypes.BOM_CREATE,
        RequestTypes.BOM_MODIFY_CHANGE
    ],

    [ActivitySheetNames.NON_M]: [
        RequestTypes.NON_M_CREATE,
    ],

    [ActivitySheetNames.SOURCE_LIST]: [
        RequestTypes.SOURCE_LIST,
    ],

    [ActivitySheetNames.MASTER_SITE]: [
        RequestTypes.SITE_CREATE,
        RequestTypes.SITE_MODIFY_CHANGE,
        RequestTypes.SITE_BLOCK,
        RequestTypes.SITE_UNBLOCK,
        RequestTypes.SITE_MODIFY_CHANGE_OTHERS
    ],

    [ActivitySheetNames.MASTER_FINANCE]: [
        RequestTypes.GL_CREATE,
        RequestTypes.GL_MODIFY_CHANGE,
        RequestTypes.GL_BLOCK,
        RequestTypes.GL_UNBLOCK,
        RequestTypes.GL_EXTEND,
        RequestTypes.COST_CENTER_CREATE,
        RequestTypes.COST_CENTER_MODIFY_CHANGE,
        RequestTypes.COST_CENTER_BLOCK,
        RequestTypes.COST_CENTER_UNBLOCK,
        RequestTypes.COST_CENTER_UNBLOCK_TEMPORARY,
        RequestTypes.COST_CENTER_UNBLOCK_BLOCK
    ],

    [ActivitySheetNames.PROFIT_CENTER]: [
        RequestTypes.PROFIT_CENTER_BLOCK,
        RequestTypes.PROFIT_CENTER_UNBLOCK
    ],

    [ActivitySheetNames.PRICING]: [
        RequestTypes.PRICING_CREATE,
        RequestTypes.PRICING_MODIFY_CHANGE
    ],

    [ActivitySheetNames.CUSTOMER]: [
        RequestTypes.SALES_PERSON_CREATE,
        RequestTypes.SALES_PERSON_MODIFY_CHANGE,
        RequestTypes.SALES_PERSON_BLOCK,
        RequestTypes.CUSTOMER_CREATE_BADAN_USAHA,
        RequestTypes.CUSTOMER_CREATE_PERORANGAN,
        RequestTypes.CUSTOMER_MODIFY_CHANGE,
        RequestTypes.CUSTOMER_EXTEND,
        RequestTypes.CUSTOMER_BLOCK,
        RequestTypes.TOP_CREDIT_LIMIT_UPDATE,
        RequestTypes.CUSTOMER_CREATE,
    ],

    [ActivitySheetNames.VENDOR]: [
        RequestTypes.BANK_KEY_MASTER_CREATE,
        RequestTypes.BANK_KEY_MASTER_MODIFY_CHANGE,
        RequestTypes.BANK_KEY_MASTER_DELETE,
        RequestTypes.VENDOR_MODIFY_CHANGE,
        RequestTypes.ACCOUNT_VENDOR_MODIFY_CHANGE,
        RequestTypes.ACCOUNT_VENDOR_ADD,
        RequestTypes.ACCOUNT_VENDOR_DELETE,
        RequestTypes.RETURN_TO_VENDOR,
        RequestTypes.VENDOR_BLOCK,
        RequestTypes.VENDOR_EXTEND,
        RequestTypes.CUSTOMER_MODIFY_CHANGE,
        RequestTypes.CUSTOMER_EXTEND
    ],
}

const ApproverConfiguration = Object.freeze({
    DEFAULT: 'DEFAULT',
    ALL: 'ALL',
})

const ATTACHMENT_SYNC_CONTEXTS = [
    {
        prop: 'REQUESTER',
        levelOrder : 0,
        constant: {
            validStatus: RequesterStatus
        }
    },
    {
        prop: 'APPROVER',
        levelOrder : 1,
        constant: {
            validStatus: ApproverStatus
        }
    },
    {
        prop: 'APPROVER_II',
        levelOrder : 2,
        constant: {
            validStatus: ApproverStatus
        }
    },
    {
        prop: 'APPROVER_III',
        levelOrder : 3,
        constant: {
            validStatus: ApproverStatus
        }
    }
];