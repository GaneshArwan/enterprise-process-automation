/**
 * Represents an archive process for managing and archiving master and child spreadsheets.
 * 
 * @param {Spreadsheet} [masterSpreadsheet=null] The master spreadsheet to be archived.
 * @param {Folder} [archiveFolder=null] The folder where the archived spreadsheets will be stored.
 */
class Archive {
    // Private static helper for result handling
    static get _Result() {
        return {
            success: (data) => ({ ok: true, data }),
            failure: (error) => ({ ok: false, error })
        }
    };

    constructor(masterSpreadsheet = null, archiveFolder = null) {
        this.masterSpreadsheet = masterSpreadsheet || getMasterSpreadsheet();
        this.archiveFolder = archiveFolder || getArchiveFolder();
    }

    // Public method to execute the archive process
    execute() {
        Logger.log('[Archive] Starting archive execution process');
        
        const archiveResult = this._archiveMasterActivity();
        if (!archiveResult.ok) {
            Logger.log(`[Archive] Master archive failed: ${archiveResult.error}`);
            console.error('Archive failed:', archiveResult.error);
            return archiveResult;
        }
        Logger.log('[Archive] Successfully completed master archive');

        // const childResult = this._processChildSpreadsheets();
        // if (!childResult.ok) {
        //     Logger.log(`[Archive] Child processing failed: ${childResult.error}`);
        //     console.error('Child processing failed:', childResult.error);
        //     return childResult;
        // }
        // Logger.log('[Archive] Successfully completed child spreadsheet processing');

        return Archive._Result.success(true);
    }

    _getTwoMonthsBefore() {
        Logger.log('[Archive] Calculating date two months before');
        const date = new Date();
        
        // Get first day of current month
        let targetDate = new Date(date.getFullYear(), date.getMonth(), 1);
        
        // Go back 2 months to get the correct archive month
        targetDate.setMonth(targetDate.getMonth() - 2);
        
        const result = {
            month: targetDate.getMonth(), // Keep 0-based for date comparison
            year: targetDate.getFullYear()
        };
        
        // For logging and display, convert month to 1-based
        Logger.log(`[Archive] Calculated date: ${result.month + 1}/${result.year}`);
        return Archive._Result.success(result);
    }
    

    _getArchivedFileName() {
        Logger.log('[Archive] Generating archived file name');
        const dateResult = this._getTwoMonthsBefore();
        if (!dateResult.ok) return dateResult;

        const { month, year } = dateResult.data;
        const fileName = `MASTER ACTIVITIES ${getMonthString(month)} ${year} ARCHIVED`;
        Logger.log(`[Archive] Generated file name: ${fileName}`);
        return Archive._Result.success(fileName);
    }

    _getRowIndicesTwoMonthsAgo(sheet) {
        Logger.log(`[Archive] Getting row indices for sheet: ${sheet.getName()}`);
        
        if (sheet.getLastRow() < ACTIVITY_HEADER_ROW_INDEX) {
            Logger.log(`[Archive] No data rows found in sheet: ${sheet.getName()}`);
            return Archive._Result.success([]);
        }

        const timestamps = getValuesByColumn(sheet, ColNames.TIMESTAMP, ACTIVITY_HEADER_ROW_INDEX);
        if (!timestamps) {
            Logger.log(`[Archive] Failed to get timestamp column for sheet: ${sheet.getName()}`);
            return Archive._Result.failure('Failed to get timestamp column');
        }

        const dateResult = this._getTwoMonthsBefore();
        if (!dateResult.ok) return dateResult;

        const { month, year } = dateResult.data;
        const matchIndices = timestamps.reduce((indices, value, idx) => {
            const date = new Date(value);
            // Compare with 0-based month from the date
            if (date.getMonth() === month && date.getFullYear() === year) {
                indices.push(idx + ACTIVITY_HEADER_ROW_INDEX);
            }
            return indices;
        }, []);

        Logger.log(`[Archive] Found ${matchIndices.length} matching rows in sheet: ${sheet.getName()}`);
        return Archive._Result.success(matchIndices);
    }

    _getExpiredRowIndices(sheet){
        const requesterValues = getValuesByColumn(
            sheet, ColNames.RESPON_REQUESTER, ACTIVITY_HEADER_ROW_INDEX);
        
        return requesterValues.reduce((indices, val, i) => {
            if (val === RequesterStatus.EXPIRED) {
                indices.push(i+ACTIVITY_HEADER_ROW_INDEX)
            }

            return indices
        }, [])
    }

    _removeRows(sheet, targetRows, keepTargetRow = false) {
        Logger.log(`[Archive] Removing rows from sheet: ${sheet.getName()}, keepTargetRow: ${keepTargetRow}`);
    
        const lastRow = sheet.getLastRow();
        const allRows = Array.from(
            { length: lastRow - ACTIVITY_HEADER_ROW_INDEX },
            (_, i) => i + ACTIVITY_HEADER_ROW_INDEX + 1
        );
    
        if (!targetRows?.length && !keepTargetRow) {
            Logger.log(`[Archive] No rows to remove from sheet: ${sheet.getName()}`);
            return Archive._Result.success(null);
        }
    
        if (!targetRows?.length && keepTargetRow) {
            Logger.log(`[Archive] Removing all non-header rows from sheet: ${sheet.getName()}`);
            
            if (lastRow > ACTIVITY_HEADER_ROW_INDEX) {
                sheet.deleteRows(ACTIVITY_HEADER_ROW_INDEX + 1, lastRow - ACTIVITY_HEADER_ROW_INDEX);
            } else {
                Logger.log("[Archive] No rows to delete.");
            }
        
            return Archive._Result.success(null);
        }
    
        const expiredRows = this._getExpiredRowIndices(sheet);
        const rowsToRemove = keepTargetRow
            ? allRows.filter(row => !targetRows.includes(row) || expiredRows.includes(row))
            : targetRows;
    
        Logger.log(`[Archive] Removing ${rowsToRemove.length} rows from sheet: ${sheet.getName()}`);
    
        if (rowsToRemove.length > 0) {
            // Sort rows in ascending order
            rowsToRemove.sort((a, b) => a - b);
    
            // Process batch deletions while dynamically adjusting for row shifts
            let adjustedShift = 0; // Keeps track of how many rows have been deleted
            let batchStart = rowsToRemove[0];
            let batchLength = 1;
    
            for (let i = 1; i < rowsToRemove.length; i++) {
                const currentRow = rowsToRemove[i];
                const previousRow = rowsToRemove[i - 1];
    
                if (currentRow === previousRow + 1) {
                    // Rows are contiguous, increase the batch length
                    batchLength++;
                } else {
                    // Rows are non-contiguous, delete the current batch
                    sheet.deleteRows(batchStart - adjustedShift, batchLength);
                    adjustedShift += batchLength; // Adjust shift for deleted rows
    
                    // Start a new batch
                    batchStart = currentRow;
                    batchLength = 1;
                }
            }
    
            // Delete the final batch
            sheet.deleteRows(batchStart - adjustedShift, batchLength);
        }
    
        return Archive._Result.success(null);
    }
    _shouldSkipSheet(sheetName) {
        // Only exclude sheets that contain '_SUBMIT'
        return sheetName.includes('_SUBMIT');
    }
    

    _makeActivityCopy() {
        Logger.log('[Archive] Creating master activity copy');
    
        if (!this.masterSpreadsheet) {
            Logger.log('[Archive] Failed to get master spreadsheet');
            return Archive._Result.failure('Failed to get master spreadsheet');
        }
    
        const fileNameResult = this._getArchivedFileName();
        if (!fileNameResult.ok) return fileNameResult;
    
        try {
            // Create new spreadsheet and move it to the archive folder
            const newSpreadsheet = SpreadsheetApp.create(fileNameResult.data);
            const newFile = DriveApp.getFileById(newSpreadsheet.getId());
            newFile.moveTo(this.archiveFolder);
    
            const masterSheets = this.masterSpreadsheet.getSheets();
            const targetSheets = newSpreadsheet.getSheets();
    
            masterSheets.forEach((sheet, index) => {
                let sheetNameMaster = sheet.getName();

                // ---> CHANGE: Only skip sheets that contain '_SUBMIT' <---
                if (!Object.values(ActivitySheetNames).includes(sheetNameMaster) || this._shouldSkipSheet(sheetNameMaster) ) {
                    Logger.log(`[Archive] Skipping sheet during copy: ${sheetNameMaster}`);
                    return; // 'return' in forEach works like 'continue'
                }
                const sourceRange = sheet.getDataRange();
                const sourceData = sourceRange.getDisplayValues();
    
                // Determine visible columns
                const visibleColumns = Array.from({ length: sourceData[0].length }, (_, col) => col)
                    .filter(col => !sheet.isColumnHiddenByUser(col + 1));
    
                // Filter data for visible columns
                const filteredData = sourceData.map(row => visibleColumns.map(col => row[col]));
    
                // Get or create target sheet
                const targetSheet = index === 0 ? targetSheets[0] : newSpreadsheet.insertSheet();
                targetSheet.setName(sheetNameMaster);
    
                // Skip if no visible data
                if (filteredData.length === 0 || filteredData[0].length === 0) return;
    
                // Batch set values
                const targetRange = targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length);
                targetRange.setValues(filteredData);
            });
    
            Logger.log(`[Archive] Successfully created archive copy: ${fileNameResult.data}`);
            return Archive._Result.success(newSpreadsheet);
        } catch (error) {
            Logger.log(`[Archive] Failed to create archive copy: ${error.message}`);
            return Archive._Result.failure(`Failed to create archive: ${error.message}`);
        }
    }
    

    _archiveMasterActivity() {
        Logger.log('[Archive] Starting master activity archival process');
        
        if (!this.masterSpreadsheet) {
            Logger.log('[Archive] Failed to get master spreadsheet');
            return Archive._Result.failure('Failed to get master spreadsheet');
        }

        const archiveResult = this._makeActivityCopy();
        if (!archiveResult.ok) return archiveResult;

        const archiveSpreadsheet = archiveResult.data;
        const archiveFile = DriveApp.getFileById(archiveSpreadsheet.getId());
        const sheetRowMap = new Map();

        try {
            // Collect rows to archive
            Logger.log('[Archive] Collecting rows to archive from master sheets');
            for (const sheet of this.masterSpreadsheet.getSheets()) {
                let sheetNameMaster = sheet.getName();
                if (!Object.values(ActivitySheetNames).includes(sheetNameMaster) || this._shouldSkipSheet(sheetNameMaster) ) {
                    Logger.log(`[Archive] Skipping invalid sheet: ${sheetNameMaster}`);
                    continue;
                }

                const rowIndicesResult = this._getRowIndicesTwoMonthsAgo(sheet);
                if (!rowIndicesResult.ok) throw new Error(rowIndicesResult.error);

                sheetRowMap.set(sheetNameMaster, rowIndicesResult.data);
            }

            // Process archive sheets
            Logger.log('[Archive] Processing archive sheets');
            for (const [sheetName, rowIndices] of sheetRowMap) {
                const archivedSheet = archiveSpreadsheet.getSheetByName(sheetName);
                if (!archivedSheet) {
                    Logger.log(`[Archive] Missing sheet in archive: ${sheetName}`);
                    throw new Error(`Missing sheet: ${sheetName}`);
                }

                const removeResult = this._removeRows(archivedSheet, rowIndices, true);
                if (!removeResult.ok) throw new Error(removeResult.error);
            }

            // Remove from master
            Logger.log('[Archive] Removing archived rows from master sheets');
            for (const [sheetName, rowIndices] of sheetRowMap) {
                const masterSheet = this.masterSpreadsheet.getSheetByName(sheetName);
                const removeResult = this._removeRows(masterSheet, rowIndices);
                if (!removeResult.ok) throw new Error(removeResult.error);
            }

            Logger.log('[Archive] Successfully completed master activity archival');
            return Archive._Result.success(true);
        } catch (error) {
            Logger.log(`[Archive] Error during master activity archival: ${error.message}`);
            try { 
                Logger.log('[Archive] Attempting to clean up failed archive file');
                archiveFile.setTrashed(true); 
            } catch { }
            return Archive._Result.failure(error.message);
        }
    }

    _processChildSpreadsheets() {
        Logger.log('[Archive] Starting child spreadsheet processing');
        
        const childUrls = getValuesByColumn(getRequestConfig(), CHILD_SPREADSHEET_KEY, 1);
        if (!childUrls) {
            Logger.log('[Archive] Failed to get child spreadsheet URLs');
            return Archive._Result.failure('Failed to get child spreadsheet URLs');
        }

        const failed = [];
        Logger.log(`[Archive] Processing ${childUrls.length} child spreadsheets`);

        for (const url of childUrls) {
            if (!url) continue;

            try {
                Logger.log(`[Archive] Processing child spreadsheet: ${url}`);
                const spreadsheet = SpreadsheetApp.openByUrl(url);
                const sheets = spreadsheet.getSheets()
                    .filter(sheet => {
                        const sheetName = sheet.getName();
                        // Only process sheets that are in ActivitySheetNames
                        return Object.values(ActivitySheetNames).includes(sheetName);
                    });

                Logger.log(`[Archive] Found ${sheets.length} activity sheets in spreadsheet: ${url}`);
                for (const sheet of sheets) {
                    const rowIndicesResult = this._getRowIndicesTwoMonthsAgo(sheet);
                    if (!rowIndicesResult.ok) {
                        Logger.log(`[Archive] Failed to get row indices for sheet: ${sheet.getName()}`);
                        failed.push(url);
                        continue;
                    }

                    const removeResult = this._removeRows(sheet, rowIndicesResult.data);
                    if (!removeResult.ok) {
                        Logger.log(`[Archive] Failed to remove rows from sheet: ${sheet.getName()}`);
                        failed.push(url);
                        break;
                    }
                }
            } catch (error) {
                Logger.log(`[Archive] Error processing spreadsheet ${url}: ${error.message}`);
                failed.push(url);
            }
        }

        if (failed.length) {
            Logger.log(`[Archive] Child processing completed with ${failed.length} failures`);
            return Archive._Result.failure(`Failed spreadsheets: ${failed.join(', ')}`);
        }

        Logger.log('[Archive] Successfully completed child spreadsheet processing');
        return Archive._Result.success(true);
    }
}