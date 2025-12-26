class AttachmentConfiguration {
    constructor() {
        this.CONFIG_CACHE = {};
        this.LOOKUP_DATA = {};
        this.loadConfiguration();
    }

    loadConfiguration() {
        const ss = SpreadsheetApp.openById(VALIDATION_CONFIGURATION_UID);
        const cfgSheet = ss.getSheetByName('Validation Config');
        const allRows = cfgSheet.getDataRange().getValues();
        const dataRows = allRows.slice(1);

        // 1) figure out needed lookup sheets
        const sheetsNeeded = new Set();
        dataRows.forEach(r => {
            const lookupSheet = r[5];
            if (lookupSheet) sheetsNeeded.add(lookupSheet);
        });

        // 2) bulk-load each sheet
        sheetsNeeded.forEach(name => {
            const sh = ss.getSheetByName(name);
            const vals = sh ? sh.getDataRange().getValues() : [];
            const hdrs = vals.length ? vals[0] : [];
            const body = vals.length ? vals.slice(1) : [];
            this.LOOKUP_DATA[name] = { headers: hdrs, body: body };
        });

        // 3) build config cache
        dataRows.forEach(r => {
            const [
                fieldName,
                validationType,
                _details,
                dataType,
                dependentField,
                lookupSheet,
                regexPattern,
                errorMessage
            ] = r;

            const cfg = {
                validationType,
                dataType,
                dependentField,
                lookupSheet,
                errorMessage: errorMessage || 'Invalid value',
                lookupSet: null,
                dependMap: null,
                regex: null
            };

            if (regexPattern) {
                cfg.regex = new RegExp(regexPattern, 'i');
            }

            // --- always build lookupSet for any lookup-based validation ---
            if (
                validationType === 'Lookup' ||
                validationType === 'Lookup Dependent' ||
                validationType === 'Lookup Regex Dependent'
            ) {
                cfg.lookupSet = this._buildLookupSet(
                    this.LOOKUP_DATA[lookupSheet],
                    fieldName
                );
            }

            // only build dependMap for dependent variants
            if (
                validationType === 'Lookup Dependent' ||
                validationType === 'Lookup Regex Dependent'
            ) {
                cfg.dependMap = this._buildDependentMap(
                    this.LOOKUP_DATA[lookupSheet],
                    fieldName,
                    dependentField
                );
            }

            this.CONFIG_CACHE[fieldName] = cfg;
        });
    }

    validate(fieldName, value, dependentValue = null) {
        const cfg = this.CONFIG_CACHE[fieldName];
        if (!cfg) return { isValid: true, message: '' };

        const v = this._normalize(value);
        const dv = this._normalize(dependentValue);
        let result;

        switch (cfg.validationType) {
            case 'Lookup':
                result = this._validateLookup(cfg, v);
                break;
            case 'Lookup Dependent':
                result = this._validateLookupDependent(cfg, v, dv);
                break;
            case 'Regex':
                result = this._validateRegex(cfg, v);
                break;
            case 'Regex Dependent':
                result = this._validateRegexDependent(cfg, v, dv);
                break;
            case 'Data Type':
                result = this._validateDataType(value, cfg.dataType);
                break;
            case 'Lookup Regex Dependent':
                result = this._validateLookupRegexDependent(cfg, v, dv);
                break;
            default:
                result = { isValid: false, message: 'Unknown validation type' };
        }

        return result;
    }

    // ——— Individual validators ———

    _validateLookup(cfg, v) {
        const ok = cfg.lookupSet.has(v);
        return { isValid: ok, message: ok ? '' : cfg.errorMessage };
    }

    _validateLookupDependent(cfg, v, dv) {
        if (!dv) return this._validateLookup(cfg, v);
        const validForDep = cfg.dependMap.get(dv) || new Set();
        const ok = validForDep.has(v);
        return { isValid: ok, message: ok ? '' : cfg.errorMessage };
    }

    _validateRegex(cfg, v) {
        const ok = cfg.regex.test(v);
        return { isValid: ok, message: ok ? '' : cfg.errorMessage };
    }

    _validateRegexDependent(cfg, v, dv) {
        const m = v.match(cfg.regex);
        const ok = m && this._normalize(m[0]) === dv;
        return { isValid: !!ok, message: ok ? '' : cfg.errorMessage };
    }

    _validateDataType(value, dataType) {
        let isValid = false, message = '';
        switch (String(dataType).toLowerCase()) {
            case 'numeric (integer)':
                isValid = /^-?\d+$/.test(value) ||
                    (typeof value === 'number' && Number.isInteger(value));
                message = isValid ? '' : 'Value must be an integer';
                break;
            case 'numeric (float)':
                isValid = !isNaN(parseFloat(value)) && isFinite(value);
                message = isValid ? '' : 'Value must be a number';
                break;
            case 'string':
                isValid = typeof value === 'string' || value instanceof String;
                message = isValid ? '' : 'Value must be text';
                break;
            default:
                message = 'Unknown data type';
        }
        return { isValid, message };
    }

    _validateLookupRegexDependent(cfg, v, dv) {
        if (!dv) return this._validateLookup(cfg, v);
        const m = dv.match(cfg.regex);
        if (!m) return { isValid: false, message: cfg.errorMessage };
        const key = this._normalize(m[1] || m[0]);
        const validForKey = cfg.dependMap.get(key) || new Set();
        const ok = validForKey.has(v);
        return { isValid: ok, message: ok ? '' : cfg.errorMessage };
    }

    // ——— Cache builders ———

    _buildLookupSet({ headers, body }, fieldName) {
        const idx = headers.indexOf(fieldName);
        const s = new Set();
        if (idx >= 0) {
            body.forEach(r => {
                String(r[idx])
                    .split('\n')
                    .map(this._normalize.bind(this))
                    .forEach(v => v && s.add(v));
            });
        }
        return s;
    }

    _buildDependentMap({ headers, body }, fieldName, dependentField) {
        const valIdx = headers.indexOf(fieldName);
        const depIdx = headers.indexOf(dependentField);
        const m = new Map();
        if (valIdx >= 0 && depIdx >= 0) {
            body.forEach(r => {
                const deps = String(r[depIdx])
                    .split('\n')
                    .map(this._normalize.bind(this));
                const vals = String(r[valIdx])
                    .split('\n')
                    .map(this._normalize.bind(this));
                deps.forEach(dep => {
                    if (!m.has(dep)) m.set(dep, new Set());
                    vals.forEach(v => v && m.get(dep).add(v));
                });
            });
        }
        return m;
    }

    _normalize(x) {
        return x == null ? '' : String(x).trim().toLowerCase();
    }
}


class AttachmentValidator {
    constructor(attachment) {
        this.attachment = attachment;
        this.attachmentConfig = new AttachmentConfiguration();
    }

    getMandatoryCells(attachmentSheet) {
        const startRow = AttachmentValues.TASK_START_ROW - 1;
        const lastColumn = attachmentSheet.getLastColumn();
        const range = attachmentSheet.getRange(startRow, 1, 1, lastColumn);
        const bgCols = range.getBackgrounds()[0];
        const hdrs = range.getValues()[0];

        const mandatoryColumns = bgCols
            .map((color, i) =>
                color === AttachmentValues.MANDATORY_COLOR
                    ? { name: hdrs[i], index: i + 1 }
                    : null
            )
            .filter(c => c);

        const numRows =
            attachmentSheet.getLastRow() - AttachmentValues.TASK_START_ROW + 1;
        if (numRows <= 0) return { columns: mandatoryColumns, values: [] };

        const values = attachmentSheet
            .getRange(
                AttachmentValues.TASK_START_ROW,
                1,
                numRows,
                lastColumn
            )
            .getValues();

        return { columns: mandatoryColumns, values };
    }

    validateCells(values, columns, attachmentSheet) {
        const sheetName = attachmentSheet.getName();
        const results = { empty: {}, invalid: {} };
        // index of the true header row
        const headerRowIndex = AttachmentValues.TASK_START_ROW - 1;

        values.forEach((row, i) => {
            // skip totally blank rows
            if (row.every(c => c == null || String(c).trim() === "")) return;

            const emptyCols = [];
            const invalidCols = [];

            columns.forEach(col => {
                const val = String(row[col.index - 1]).trim();
                const cfg = this.attachmentConfig.CONFIG_CACHE[col.name];

                // --- lookup the dependent-field value against the full header ---
                let dep = null;
                if (cfg?.dependentField) {
                    const depColIdx = getColumnIndex(
                        attachmentSheet,
                        cfg.dependentField,
                        headerRowIndex
                    );
                    if (depColIdx > 0) {
                        dep = this.attachmentConfig._normalize(row[depColIdx - 1]);
                    }
                }

                if (!val) {
                    emptyCols.push(col.name);
                } else {
                    const { isValid, message } = this.attachmentConfig.validate(
                        col.name,
                        val,
                        dep
                    );
                    if (!isValid) invalidCols.push({ colName: col.name, message });
                }
            });

            const rowNum = i + AttachmentValues.TASK_START_ROW;
            if (emptyCols.length) {
                results.empty[sheetName] = results.empty[sheetName] || {};
                results.empty[sheetName][rowNum] = emptyCols;
            }
            if (invalidCols.length) {
                results.invalid[sheetName] = results.invalid[sheetName] || {};
                results.invalid[sheetName][rowNum] = invalidCols;
            }
        });

        return results;
    }

    generateSummary(validationResults) {
        if (!validationResults || (Object.keys(validationResults.EMPTY_MANDATORY).length === 0 && Object.keys(validationResults.INVALID_VALUES).length === 0)) {
            return "✅ All attachments have passed the validation checks.";
        }

        let summary = [];

        // Header
        summary.push("<b>Action Required:</b><br><br>");

        // Count total issues
        const totalEmpty = Object.values(validationResults.EMPTY_MANDATORY).reduce((sum, rows) => sum + Object.keys(rows).length, 0);
        const totalInvalid = Object.values(validationResults.INVALID_VALUES).reduce((sum, rows) => sum + Object.keys(rows).length, 0);

        // Overview of issues
        if (totalEmpty > 0) {
            summary.push(`• <b>${totalEmpty}</b> row(s) require missing values to be filled.<br>`);
        }

        if (totalInvalid > 0) {
            summary.push(`• <b>${totalInvalid}</b> row(s) contain invalid values that need correction.<br>`);
        }

        return summary.join("");
    }

    clearRemarks(attachmentSheet) {
        const remarksCol = getColumnIndex(
            attachmentSheet,
            "Remarks(MDM)",
            AttachmentValues.TASK_START_ROW - 1
        );
        if (remarksCol < 1) return;

        const lastRow = attachmentSheet.getLastRow();
        const numRows = lastRow - AttachmentValues.TASK_START_ROW + 1;
        if (numRows > 0) {
            attachmentSheet
                .getRange(
                    AttachmentValues.TASK_START_ROW,
                    remarksCol,
                    numRows,
                    1
                )
                .clearContent();
        }
    }

    handleValidationIssues(attachmentSheet, issuesByRow, type) {
        const sheetName = attachmentSheet.getName();
        const sheetIssues = issuesByRow[sheetName];
        if (!sheetIssues) return;

        const startRow = AttachmentValues.TASK_START_ROW;
        const lastRow = attachmentSheet.getLastRow();
        const numRows = lastRow - startRow + 1;
        if (numRows <= 0) return;

        const remarksCol = getColumnIndex(
            attachmentSheet,
            "Remarks(MDM)",
            startRow - 1
        );
        const statusCol = getColumnIndex(
            attachmentSheet,
            "Status (MDM)",
            startRow - 1
        );
        if (remarksCol < 1 || statusCol < 1) return;

        const remarksRange = attachmentSheet.getRange(
            startRow,
            remarksCol,
            numRows,
            1
        );
        const statusRange = attachmentSheet.getRange(
            startRow,
            statusCol,
            numRows,
            1
        );
        const remarksVals = remarksRange.getValues();
        const statusVals = statusRange.getValues();

        Object.entries(sheetIssues).forEach(([r, details]) => {
            const idx = parseInt(r, 10) - startRow;
            let newRemark;

            if (type === "Empty Mandatory") {
                newRemark = `- ${type}: ${details.join(", ")}`;
            } else {
                newRemark = details
                    .map(d => `- ${d.colName}: (${d.message})`)
                    .join("\n");
            }

            const existing = String(remarksVals[idx][0] || "").trim();
            remarksVals[idx][0] = existing
                ? `${existing}\n\n${newRemark}`
                : newRemark;
            // statusVals[idx][0] = 'Rejected';
        });

        remarksRange.setValues(remarksVals);
        statusRange.setValues(statusVals);
    }

    execute() {
        const attachment = this.attachment.getAttachment();
        if (!attachment || attachment.getName().includes("PROMO")) return {};

        const results = { EMPTY_MANDATORY: {}, INVALID_VALUES: {} };

        for (const sheet of attachment.getSheets()) {
            if (sheet.getTabColor() !== TASK_SHEET_COLOR) continue;

            const { values, columns } = this.getMandatoryCells(sheet);
            const validation = this.validateCells(values, columns, sheet);

            this.clearRemarks(sheet);
            if (Object.keys(validation.empty).length) {
                this.handleValidationIssues(sheet, validation.empty, "Empty Mandatory");
                Object.assign(results.EMPTY_MANDATORY, validation.empty);
            }
            if (Object.keys(validation.invalid).length) {
                this.handleValidationIssues(sheet, validation.invalid, "Invalid Values");
                Object.assign(results.INVALID_VALUES, validation.invalid);
            }
        }

        return results;
    }
}
