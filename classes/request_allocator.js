class RequestAllocator {
  constructor(request) {
    this.request            = request;
    this.activity           = request.activity;
    this.masterConfig       = new MasterConfig();
    // SANITIZED: Use generic default agent
    this.DEFAULT_ALLOCATION = 'AGENT_01';
    this.LOG_PREFIX         = '[RequestAllocator]';
  }

  _log(level, ...msgs) {
    const text = `${this.LOG_PREFIX} ${msgs.join(' ')}`;
    switch(level) {
      case 'info':  console.info(text);  break;
      case 'warn':  console.warn(text);  break;
      case 'error': console.error(text); break;
      default:      console.log(text);   break;
    }
  }

  /**
   * Convert "HH:MM:SS" (or "MM:SS") into the specified unit.
   * unit = 'seconds' | 'minutes' | 'hours'
   * Returns Infinity on invalid input.
   */
  _convertTime(str, unit = 'seconds') {
    const secs = parseHms(str);
    if (!isFinite(secs)) return Infinity;

    switch (unit) {
      case 'minutes':
        return secs / 60;
      case 'hours':
        return secs / 3600;
      case 'seconds':
      default:
        return secs;
    }
  }

  _getWorkload(name) {
    const cleanName = String(name || '').trim();
    
    // 1. Get STATUS from Sheet
    const sheet = getMasterSpreadsheet('WORKLOAD MDM');
    const [rawNames, statuses] = getValuesByColumns(sheet, ['MDM Name', 'Status'], 1);

    const normalizedSheetNames = rawNames.map(n => String(n || '').trim().toLowerCase());
    const idx = normalizedSheetNames.indexOf(cleanName.toLowerCase());

    // Default status if name not found: TRUE (Assume busy/inactive)
    let status = true; 
    let realName = cleanName;

    if (idx !== -1) {
       const rawStatus = statuses[idx];
       status = !!String(rawStatus || '').trim(); 
       realName = rawNames[idx]; 
    } else {
       this._log('warn', `MDM "${name}" not found in sheet for Status check.`);
    }

    // 2. Get TOTAL TIME from Properties
    const totalTimeSeconds = getMdmWorkloadFromProperty(realName);

    const wl = {
      mdmName: realName,
      totalTimeSeconds: totalTimeSeconds, 
      status: status 
    };

    this._log('info', `Workload "${wl.mdmName}": Status=${wl.status}, TotalTime=${wl.totalTimeSeconds}s (via Props)`);
    return wl;
  }

  isSpecialRequest(dep) {
    return dep === 'SPECIAL PROJECT';
  }

  updateMdmWorkload(mdmName, timeToAddSeconds) {
    updateMdmWorkloadProperty(mdmName, timeToAddSeconds);
  }

  allocate() {
    const m = this.activity.getActivityValueMap();
    const currentRequestType = m.REQUEST_TYPE;

    // "Special Project" check
    if (this.isSpecialRequest(m.DEPARTMENT)) { 
      this._log('info', `Special project → default "${this.DEFAULT_ALLOCATION}"`);
      return this.DEFAULT_ALLOCATION;
    }

    // --- UNIFIED ALLOCATION LOGIC ---

    // 1. Get Permission Matrix from "Distribution" Sheet
    const matrix = getMdmDistributionMatrix(); 
    
    let matrixCandidates = [];
    if (matrix && matrix[currentRequestType]) {
        matrixCandidates = matrix[currentRequestType];
    }

    // IF IN MATRIX -> Run Scaled RR
    if (matrixCandidates.length > 0) {
        const candidateNames = matrixCandidates.map(name => name.toUpperCase());
        this._log('info', `🎯 Matrix Match (Distribusi) for "${currentRequestType}". Candidates: [${candidateNames.join(', ')}].`);

        // A. Get Workload Data
        const candidatesData = candidateNames
            .map(name => this._getWorkload(name)) 
            .filter(data => data !== null);

        // B. Filter only AVAILABLE (Status = FALSE)
        const availableCandidates = candidatesData.filter(c => c.status === false);

        if (availableCandidates.length > 0) {
             // C. Find Lowest Workload
             const minWorkload = Math.min(...availableCandidates.map(c => c.totalTimeSeconds));
             
             // D. Get list of candidates with lowest workload
             const bestCandidates = availableCandidates
                 .filter(c => c.totalTimeSeconds === minWorkload)
                 .map(c => c.mdmName);

             this._log('debug', `Lowest Workload (Total Time): ${minWorkload}s. Finalists: [${bestCandidates.join(', ')}]`);

             // E. Pick Winner (Round-Robin if tie)
             let assignedMdm = null;
             if (bestCandidates.length === 1) {
                 assignedMdm = bestCandidates[0];
                 this._log('info', `Scaled RR: Assigning to ${assignedMdm} (Winner - Lowest Load).`);
             } else {
                 const ruleKey = `MATRIX_RR|${currentRequestType}`; 
                 assignedMdm = getNextMdmViaRoundRobin(ruleKey, bestCandidates);
                 this._log('info', `Scaled RR: Tie-break used. Selected: ${assignedMdm}.`);
             }

             if (assignedMdm) return assignedMdm;

        } else {
            this._log('warn', `Scaled RR: All candidates busy (Status=TRUE).`);
        }
        this._log('warn', `Scaled RR failed, falling back to BAU.`);
    } else {
        this._log('info', `Request Type "${currentRequestType}" not found in Distribution Matrix. Using BAU.`);
    }

    // --- BAU LOGIC (FALLBACK) ---
    this._log('info', `⚙️ Executing BAU allocation for "${currentRequestType}".`);

    const allocationRule = this.masterConfig.getWorkAllocation({
        companyCode: m.COMPANY_CODE_NAME,
        requestType: m.REQUEST_TYPE,
        department: m.DEPARTMENT,
    }); 

    if (!allocationRule) {
      this._log('warn', `BAU: No allocation rule found... → default "${this.DEFAULT_ALLOCATION}"`);
      return this.DEFAULT_ALLOCATION;
    }

    const candidates = [allocationRule.pic, ...(allocationRule.backups || [])];
    for (const raw of candidates) {
        const names = raw
        .split(',')
        .map(n => n.trim())
        .filter(n => n);

        if (names.length === 0) continue;

        const workloads = names
        .map(n => this._getWorkload(n))
        .filter(wl => wl !== null);

        if (workloads.length === 0) {
          this._log('warn', `No valid workloads for "${raw}", skipping`);
          continue;
        }
        // if *all* are busy, skip
      if (workloads.every(wl => wl.status === true)) {
        this._log(
          'info',
          `All of [${names.join(', ')}] busy → skipping "${raw}"`
        );
        continue;
      }

      // pick only the free ones
      const freeOnes = workloads.filter(wl => wl.status === false);
      // choose the one with the smallest seconds
      const best = freeOnes.reduce((a, b) =>
        b.totalTimeSeconds < a.totalTimeSeconds ? b : a
      );

      this._log(
        'info',
        `Allocating to "${best.mdmName}" (from "${raw}") with lowest ` +
        `Est.Time=${best.totalTimeSeconds}s`
      );
      return best.mdmName;
    }

    this._log(
      'warn',
      `No available PIC/backups → default "${this.DEFAULT_ALLOCATION}"`
    );
    return this.DEFAULT_ALLOCATION;
  }
}