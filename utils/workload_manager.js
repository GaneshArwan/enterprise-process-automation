function getMdmWorkloadFromProperty(mdmName) {
  if (!mdmName) return 0;
  const key = `WORKLOAD_TIME_${mdmName.toUpperCase().trim()}`;
  const props = PropertiesService.getScriptProperties();
  const val = props.getProperty(key);
  return Number(val) || 0; 
}

function updateMdmWorkloadProperty(mdmName, secondsToAdd) {
    if (!mdmName || secondsToAdd === undefined) return 0;

    const key = `WORKLOAD_TIME_${mdmName.toUpperCase().trim()}`;
    const lock = LockService.getScriptLock();

    try {
        lock.waitLock(10000); // Wait up to 10 seconds
    } catch (e) {
        Logger.log(`[WorkloadManager] Lock timeout for ${mdmName}: ${e.message}`);
        throw new Error("Server busy, please try again.");
    }

    let newTotal = 0;
    try {
        const props = PropertiesService.getScriptProperties();
        const currentVal = Number(props.getProperty(key)) || 0;

        newTotal = currentVal + secondsToAdd;
        if (newTotal < 0) newTotal = 0; // Prevent negative time

        props.setProperty(key, String(newTotal));
        Logger.log(`[WorkloadManager] Updated ${mdmName}: ${currentVal}s -> ${newTotal}s`);

    } catch (e) {
        Logger.log(`[WorkloadManager] Error: ${e.message}`);
        throw e;
    } finally {
        lock.releaseLock();
    }

    return newTotal;
}

function syncSheetToProperties() {
  const sheet = getMasterSpreadsheet('WORKLOAD MDM');
  const data = sheet.getDataRange().getValues();
  // Asumsi: Col A=Name, Col E=Total Time (Sesuaikan index kolom Anda)
  // Header row 1
  
  const props = PropertiesService.getScriptProperties();
  const updates = {};

  // Loop data sheet (mulai baris 2)
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim().toUpperCase(); // Col A (Index 0)
    const timeVal = data[i][4]; // Col E (Index 4) - Pastikan ini kolom Total Time
    
    let seconds = 0;
    if (typeof timeVal === 'number') seconds = timeVal;
    else if (String(timeVal).includes(':')) seconds = parseHms(timeVal); // pastikan parseHms tersedia
    else seconds = parseInt(timeVal) || 0;

    if (name) {
      const key = `WORKLOAD_TIME_${name}`;
      updates[key] = String(seconds);
    }
  }
  
  props.setProperties(updates);
  Logger.log("Migrasi Selesai: Data Sheet berhasil disalin ke Script Properties.");
}