class RequestLogger {
  constructor(request) {
    this.request            = request;
    this.activity           = request.activity;
    this.masterConfig       = new MasterConfig();
    this.logSheet           = getMasterSpreadsheet('REQUEST LOG')
  }

  addMasterLog(activity, actor, reason = null){
    const logSheet = getMasterSpreadsheet('MASTER LOG');
    const {REQUEST_NUMBER} = this.activity.getActivityValueMap();

    const entry = [
      REQUEST_NUMBER,
      getDateNow(),
      activity,
      actor,
      reason || ''
    ];

    logSheet.appendRow(entry);
    Logger.log(`Logged for Request Number ${REQUEST_NUMBER}`);
  }

  addWorkspaceLog(modification, oldValue, newValue){
    const logSheet = getMasterSpreadsheet('WORKSPACE LOG');
    const {REQUEST_NUMBER, COMPANY_CODE_NAME, DEPARTMENT, REQUEST_TYPE} = this.activity.getActivityValueMap();

    const modifiedBy = Session.getEffectiveUser().getEmail();
    const timestamp  = getDateNow();

    const logEntry = [
      REQUEST_NUMBER,
      timestamp,
      modification,
      COMPANY_CODE_NAME,
      DEPARTMENT || '',
      REQUEST_TYPE,
      oldValue,
      newValue,
      modifiedBy
    ];

    logSheet.appendRow(logEntry);
    Logger.log(`Logged change for Request Number ${REQUEST_NUMBER}`);
  }
}