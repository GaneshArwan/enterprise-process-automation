function hasKeywords(value, keywords) {
    if (!value) return false;
    if (!Array.isArray(keywords)) keywords = [keywords]; // Convert to array if it's a single string
    return keywords.some(
        keyword => value.toString().startsWith(keyword)
    );
}

function isUrl(value) {
    return typeof value === 'string' && value.startsWith('https://')
}

function isFolderUrl(url) {
  return typeof url === 'string' && url.includes('/folders/');
}

function isString(value) {
    return typeof value === 'string' || value instanceof String;
}

function isNumber(value) {
    return typeof value === 'number' || value instanceof Number;
}

function extractCompanyCode(value) {
    return value.toString().slice(0, 4).trim();
}

function extractCompanyName(value) {
    if (value.includes('-')) {
        return value.toString().slice(6).trim();
    }
    return value;
}

function getCompanyFullName(value) {
    return COMPANY_NAME_MAP[value];
}

function extractPromoCode(value) {
    return value.substring(0, 4);
}

function extractAccessSequence(value) {
    return value.split('(')[0].trim();
}

function upperAndSeparate(value, separateBy = '_') {
    upperCaseVal = value.toString().toUpperCase().trim();
    return upperCaseVal.replace(/[\s-]+/g, separateBy)
}

function isNotEmpty(value) {
    return (typeof value === "string" ? value.trim() : value) !== "" &&
        value !== null &&
        value !== undefined;
}

function extractSheetId(value) {
    const pattern = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = value.match(pattern);

    if (match) return match[1];
    return null;
}

function extractDriveId(value) {
    const pattern = /[\?&]id=([a-zA-Z0-9-_]+)/;
    const match = value.match(pattern);

    if (match) return match[1];
    return null;
}

function extractFileIdFromUrl(url) {
    const idMatch = url.match(/[-\w]{25,}/);
    if (idMatch && idMatch[0]) {
        return idMatch[0];
    }
    throw new Error('Invalid file URL: ' + url);
}

function containsEmptyString(values) {
    return values.some(row => row.includes(""));
}

function removeSubmitSuffix(value) {
    return value.replace(/_SUBMIT$/, '');
}