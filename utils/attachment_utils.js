function getMandatoryColumns(attachmentSheet){
    const startRow = AttachmentValues.TASK_START_ROW - 1;
    const lastColumn = attachmentSheet.getLastColumn();

    const range = attachmentSheet.getRange(startRow, 1, 1, lastColumn)
    const backgroundCols = range.getBackgrounds()[0];

    const mandatoryColumnIndices = backgroundCols
        .map((color, index) => (
            color === AttachmentValues.MANDATORY_COLOR 
            ? index + 1 
            : null)) 
        .filter(index => index !== null); 

    return mandatoryColumnIndices; 
}

function getMandatoryValues(attachmentSheet){
    const mandatoryColumns = getMandatoryColumns(attachmentSheet);
    const numRows = attachmentSheet.getLastRow() - AttachmentValues.TASK_START_ROW + 1;
    if (numRows <= 0) return {}
    
    // Get all values starting from AttachmentValues.TASK_START_ROW to the last row
    const allValues = attachmentSheet.getRange(
        AttachmentValues.TASK_START_ROW, 1, 
        numRows, attachmentSheet.getLastColumn()
    ).getValues();

    // Extract only the values in mandatory columns for each row
    const mandatoryValues = allValues.map(row => 
        mandatoryColumns.map(colIndex => row[colIndex - 1])
    );


    return {
        columns: mandatoryColumns,
        values: mandatoryValues
    };
}