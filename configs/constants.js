// Development Environment Constants
const DEV_CONSTANTS = {
    EMAIL_MDM_GROUP: 'admin.dev@example.com',
    EMAIL_CC_BOM: 'manager.bom@example.com',
    EMAIL_CC_SECOND: 'manager.second@example.com',
    // REPLACE THESE WITH YOUR OWN DUMMY SHEET IDs FOR THE DEMO
    MASTER_CONFIGURATION_UID: "REPLACE_WITH_MASTER_SHEET_ID",
    VALIDATION_CONFIGURATION_UID: "REPLACE_WITH_VALIDATION_SHEET_ID",
    APPROVER_CONFIGURATION_UID: "REPLACE_WITH_APPROVER_SHEET_ID",
    ARCHIVE_FOLDER_ID: 'REPLACE_WITH_FOLDER_ID',
    REQUEST_FOLDER_ID: 'REPLACE_WITH_FOLDER_ID',
    MDM_WORKSPACE_ID: 'REPLACE_WITH_WORKSPACE_ID',
    LOG_SPREADSHEET_ID :'REPLACE_WITH_LOG_ID',
    WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec'
}

// Production Environment Constants
const PROD_CONSTANTS = {    
    EMAIL_MDM_GROUP: 'mdm.group@example-corp.com',
    EMAIL_CC_BOM: 'dept.bom@example-corp.com',
    EMAIL_CC_SECOND: 'dept.second@example-corp.com',
    MASTER_CONFIGURATION_UID: 'REPLACE_WITH_PROD_MASTER_ID',
    VALIDATION_CONFIGURATION_UID: "REPLACE_WITH_PROD_VALIDATION_ID",
    APPROVER_CONFIGURATION_UID: "REPLACE_WITH_PROD_APPROVER_ID",
    ARCHIVE_FOLDER_ID : 'REPLACE_WITH_PROD_FOLDER_ID',
    REQUEST_FOLDER_ID: 'REPLACE_WITH_PROD_FOLDER_ID',
    MDM_WORKSPACE_ID: 'REPLACE_WITH_PROD_WORKSPACE_ID',
    LOG_SPREADSHEET_ID :'REPLACE_WITH_PROD_LOG_ID', 
    WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_PROD_DEPLOYMENT_ID/exec'
}

/**
 * Determine which set of constants to use based on environment
 * Set values of constants based on environment
 * **/
const service = PropertiesService.getScriptProperties().getProperties();
let constants = (service.ENVIRONMENT == 'DEVELOPMENT') ? DEV_CONSTANTS : PROD_CONSTANTS;

const SUBMIT_SUFFIX = '_SUBMIT';
const SYSTEM_SENT_BACK_SEPARATOR = "\n--\n";
const ACTIVITY_HEADER_ROW_INDEX = 2;

// Mapping sheet names to abbreviations
const SHEET_ABBR_MAP = {
    "EXTEND PIR": "EXTPIR",
    "BOM": "BOM",
    "PROMOTION": "PROMO",
    "BASIC DATA": "BD",
    "HIERARCHY": "AH",
    "SOURCE LIST": "SCLIST",
    "NON M": "NONM",
    "MERCHANDISE": "MERCH",
    "STATUS/LISTING": "STSLST",
    "MASTER DATA": "MSTRDATA",
    "IMAGE": "IMG",
    "MASTER FINANCE": "FINANCE",
    "MASTER SITE" : "SITE",
    "PRICING": "PRICING",
    "PROFIT CENTER": "PFTCTR",
    "CUSTOMER": "CST",
    "VENDOR": 'VNDR',
}

// Approver Information
const NO_APPROVER = 'NO APPROVER';
const VALID_APPROVER_STATUS = [
    "Approved", 
    "Partially Rejected",
    "Rejected"
];

// Email Configuration
const EMAIL_MDM_GROUP = constants.EMAIL_MDM_GROUP
const EMAIL_CC_BOM = constants.EMAIL_CC_BOM;
const EMAIL_CC_SECOND = constants.EMAIL_CC_SECOND;


// Date Configuration
const TASK_SHEET_COLOR = '#6d9eeb';
const EXPIRED_DAY_LIMIT = 3;

// Master Configuration
const DRIVE_SUFFIX = '_DRIVE';
const IMAGE_DRIVE_SUFFIX = '_IMAGE';
const CHILD_SPREADSHEET_KEY = 'SPREADSHEET';
const MASTER_CONFIGURATION_UID = constants.MASTER_CONFIGURATION_UID;
const ARCHIVE_FOLDER_ID = constants.ARCHIVE_FOLDER_ID;
const VALIDATION_CONFIGURATION_UID = constants.VALIDATION_CONFIGURATION_UID;
const APPROVER_CONFIGURATION_UID = constants.APPROVER_CONFIGURATION_UID;
const REQUEST_FOLDER_ID = constants.REQUEST_FOLDER_ID;
const MDM_WORKSPACE_ID = constants.MDM_WORKSPACE_ID;
const LOG_SPREADSHEET_ID = constants.LOG_SPREADSHEET_ID;    
const WEB_APP_URL = constants.WEB_APP_URL;

// Default Value Maps
const REQUEST_TYPE_DEFAULT_VALUE_MAP = {
    'BASIC DATA': 'Basic Data Modify/Change',
    'NON M': 'Non M Create',
    'SOURCE LIST': 'Source List',
    'MERCHANDISE': 'Merchandise Create',
    'IMAGE': 'Image Modify/Change'
};

const DEPARTMENT_DEFAULT_VALUE_MAP = {
    'NON M': 'NON MERCHANDISE',
    'MASTER FINANCE': 'GENERAL LEDGER',
    'PROFIT CENTER': 'FINANCE CONTROL'
}

// Mapped Company Codes to Generic Business Unit Names
const COMPANY_NAME_MAP = {
    "BU01": "Retail Unit Alpha",
    "BU02": "Retail Unit Beta",
    "BU03": "Home Essentials",
    "BU04": "Home Essentials Retail",
    "BU05": "Industrial Solutions A",
    "BU06": "Industrial Solutions B",
    "BU07": "Machinery Corp",
    "BU08": "Toys & Games",
    "BU09": "Innovative Tech",
    "BU10": "Sensor Systems",
    "BU11": "Tooling Depot",
    "BU12": "Creative Solutions",
    "BU13": "Pet Care Unit",
    "BU14": "Rental Services",
    "BU15": "Integrated Services",
    "BU16": "Property Management",
    "BU17": "Lifestyle Goods",
    "BU18": "Digital Systems",
    "BU19": "Global Tools",
    "BU20": "Food & Beverage A",
    "BU21": "Finance Unit",
    "CORP": "Corporate Headquarters"
};

const MENU_LINK = {
    "MDM Toolkit": "https://example.com/toolkit",
    "Dashboard": "https://example.com/dashboard",
    "MDM Knowledge Center": "https://example.com/docs"
};