function getDriveFolder(sheetName, companyName, withImageFolder = false) {
  const activity = upperAndSeparate(sheetName);
  const config = [
    ActivitySheetNames.MASTER_FINANCE,
    ActivitySheetNames.MASTER_SITE,
    ActivitySheetNames.PRICING,
    ActivitySheetNames.PROFIT_CENTER,
    ActivitySheetNames.CUSTOMER,
    ActivitySheetNames.VENDOR
  ].includes(sheetName)
    ? getRequestConfig(sheetName)
    : getRequestConfig(companyName)

  const requestDriveID = config[activity + MasterConfiguration.DRIVE_SUFFIX];
  const requestDrive = DriveApp.getFolderById(requestDriveID);

  const additionalDriveID = config[activity + MasterConfiguration.ADDITIONAL_SUFFIX]
  let additionalDrive
  if (additionalDriveID) {
    additionalDrive = DriveApp.getFolderById(additionalDriveID);
  }

  let imageDrive = null;
  if (withImageFolder) {
    const imageDriveID = config[activity + MasterConfiguration.IMAGE_DRIVE_SUFFIX];
    imageDrive = DriveApp.getFolderById(imageDriveID);
  }

  return {
    requestDrive: requestDrive,
    additionalDrive: additionalDrive,
    imageDrive: imageDrive
  }
}

function createTxtFile(name, text, driveFolder) {
  const txtFile = driveFolder.createFile(name, text)
  return txtFile;
}

function addDriveEditors(fileId, emails) {
  if (!fileId || !emails) {
    Logger.log(`File ID : ${fileId} or emails :${emails} are not provided. Cannot add editors.`);
    return;
  }

  // Normalize and dedupe
  const list = Array.from(new Set((emails || []).map(e => (e || '').trim()).filter(Boolean)));
  if (list.length === 0) return;

  try {
    // If this is a Sheet ID, use SpreadsheetApp for batch add; otherwise fall back
    const ss = SpreadsheetApp.openById(fileId);
    ss.addEditors(list);
    Logger.log(`[addDriveEditors] Added ${list.length} editors via batch to ${fileId}`);
    return;
  } catch (_) { /* not a spreadsheet or access limitations */ }

  // Fallback: chunked Drive API calls with small sleep to mitigate rate limits
  const chunkSize = 50;
  for (let i = 0; i < list.length; i += chunkSize) {
    const chunk = list.slice(i, i + chunkSize);
    for (const email of chunk) {
      try {
        Drive.Permissions.insert(
          { role: 'writer', type: 'user', value: email },
          fileId,
          { sendNotificationEmails: 'false' }
        );
      } catch (err) {
        try {
          Drive.Permissions.insert(
            { role: 'writer', type: 'user', value: email },
            fileId,
            { sendNotificationEmails: 'true' }
          );
        } catch (err2) {
          Logger.log(`[addDriveEditors] Failed to add ${email}. Error: ${err2 && err2.message}`);
        }
      }
    }
    Utilities.sleep(200);
  }
}
