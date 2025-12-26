/**
 * Retrieves all request rows that:
 * - Respon Requester = "Completed"
 * - All approvers are "Approved", "Partially Rejected", or empty
 * - Processed By is still empty
 *
 * @param {Sheet} sheet
 * @return {Object[]} Array of row objects ready for processing
 */
function getUnprocessedApprovedRequests(sheet) {
  if (!sheet) {
    Logger.log("❌ Sheet object is undefined/null.");
    return [];
  }

  // Get all data
  const data = sheet.getDataRange().getValues();
  if (data.length <= 2) {
    Logger.log(`⚠️ Sheet "${sheet.getName()}" has no data rows.`);
    return [];
  }

  // Use row 2 as headers
  const headers = data[1];
  const rows = data.slice(2); // Data starts from row 3

  const approverCols = headers.filter(h => h && h.startsWith("RESPON_"));
  const processedByIndex = headers.indexOf("Processed By");
  const responRequesterIndex = headers.indexOf("Respon Requester");

  if (processedByIndex === -1 || responRequesterIndex === -1) {
    Logger.log(`⚠️ Sheet "${sheet.getName()}" missing necessary columns.`);
    return [];
  }

  const result = [];

  rows.forEach((row, i) => {
    const rowIndex = i + 3; // +3 because of headers in row 2

    // Step 1: Check Respon Requester
    const responRequester = row[responRequesterIndex];
    if (responRequester !== "Completed") {
      Logger.log(`🔸 Row ${rowIndex} skipped (Respon Requester is "${responRequester}")`);
      return;
    }

    // Step 2: Check PROCESSED_BY
    const processedBy = row[processedByIndex];
    if (processedBy && processedBy !== "") {
      Logger.log(`🔹 Row ${rowIndex} skipped (already has Processed By: ${processedBy})`);
      return;
    }

    // Step 3: Check approvals
    const approvals = approverCols.map(col => {
      const idx = headers.indexOf(col);
      return row[idx];
    });

    const isNoApprover = approvals.every(v => !v || v === "");
    const isAllApprovedOrPartial = approvals.every(v => 
      v === "Approved" || 
      v === "Partially Rejected" || 
      v === ""
    );

    if (!(isNoApprover || isAllApprovedOrPartial)) {
      Logger.log(`🟡 Row ${rowIndex} skipped (not fully approved/partially rejected). Approvals: ${JSON.stringify(approvals)}`);
      return;
    }

    Logger.log(`✅ Row ${rowIndex} is ready for allocation.`);

    result.push({
      rowIndex,
      approvals,
      responRequester,
      processedBy
    });
  });

  return result;
}

/**
 * Processes all requests that have been approved by all approvers
 * (or have no approvers) and have not yet been allocated
 * across all specified sheets.
 */
function processPendingAllocations() {
  const sheetNames = [
    "EXTEND PIR",
    "MERCHANDISE", 
    "MASTER DATA", 
    "STATUS/LISTING", 
    "BASIC DATA", 
    "PROMOTION",
    "IMAGE",
    "HIERARCHY",
    "BOM",
    "NON M",
    "SOURCE LIST",
    // "MASTER SITE",
    // "MASTER FINANCE",
    // "VENDOR", 
    // "CUSTOMER", 
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`❌ Sheet "${sheetName}" not found, skipping.`);
      return;
    }

    Logger.log(`🔍 Checking sheet: "${sheetName}"`);

    const pendingRequests = getUnprocessedApprovedRequests(sheet);

    if (pendingRequests.length === 0) {
      Logger.log(`📭 No valid requests found in sheet "${sheetName}".`);
      return;
    }

    Logger.log(`📦 Sheet "${sheetName}": ${pendingRequests.length} request(s) ready for allocation.`);

    pendingRequests.forEach(r => {
      Logger.log(`➡️ Processing row ${r.rowIndex} in sheet "${sheetName}".`);

      const request = new Request(sheet, r.rowIndex);
      const handler = new RequestHandler(request);

      const processedBy = handler.handleRequestApproved();

      if (processedBy) {
        Logger.log(`✅ Row ${r.rowIndex} allocated to: ${processedBy}`);
      } else {
        Logger.log(`❌ Failed to allocate row ${r.rowIndex}`);
      }
    });
  });
}
