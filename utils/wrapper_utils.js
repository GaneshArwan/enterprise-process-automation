const BACKOFF_CFG = {
  1: { base: 50, max: 250 },   // high priority → retry fast
  2: { base: 100, max: 500 },   // normal
  3: { base: 200, max: 1000 },  // background
  4: { base: 300, max: 1500 },
  5: { base: 400, max: 2000 }
};

const GUARD_WAIT_MS = 200;
const LOCK_TIMEOUT_MS = 300000;
const STALE_THRESHOLD_MS = 8000;
const TTL_CUSHION_MS = 2000;

/**************************************
 * wrapper_utils.js (additions)
 **************************************/

const KEY_LOCK_PREFIX = 'KEY_LOCK_';

function makeKeyLockKey(lockKey) {
  const ns = 'KEY_LOCK_v2:';
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    ns + String(lockKey)
  );
  const hex = digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  return `${ns}${hex}`; // fixed-length, no collisions in practice
}
/**
 * Acquire a general-purpose distributed lock keyed by arbitrary string.
 * Same semantics as acquireSheetLockWithPriority (no preemption of healthy locks).
 */
function acquireKeyLock(lockKey, operation, priority = 2, maxWaitTime = 60000) {
  const cache = CacheService.getScriptCache();
  const key = makeKeyLockKey(lockKey);
  const start = Date.now();
  const myId = Utilities.getUuid();
  let attempt = 0;

  Logger.log(`[${operation}] Attempting key lock "${lockKey}" (prio ${priority})`);

  while (Date.now() - start < maxWaitTime) {
    attempt += 1;
    const guard = LockService.getScriptLock();
    try {
      guard.waitLock(GUARD_WAIT_MS);

      const now = Date.now();
      const raw = cache.get(key);
      let takeOver = false;
      let existing = null;

      if (!raw) {
        takeOver = true;
      } else {
        try {
          existing = JSON.parse(raw);
          const lastBeat = existing.lastHeartbeat || existing.timestamp || 0;
          const beatAge = now - lastBeat;
          const isExpired = now > (existing.expiry || 0);

          if (isExpired || beatAge > STALE_THRESHOLD_MS) {
            Logger.log(`[${operation}] Stale/expired key lock detected; taking over.`);
            takeOver = true;
          } else {
            Logger.log(`[${operation}] Key lock held by "${existing.operation}" (prio ${existing.priority}); waiting.`);
          }
        } catch (e) {
          Logger.log(`[${operation}] Malformed key lock data (${e}); taking over.`);
          takeOver = true;
        }
      }

      if (takeOver) {
        const now2 = Date.now();
        const payloadObj = {
          operation,
          timestamp: now2,
          lastHeartbeat: now2,
          priority,
          lockId: myId,
          expiry: now2 + LOCK_TIMEOUT_MS
        };
        const ttlSec = Math.ceil((LOCK_TIMEOUT_MS + TTL_CUSHION_MS) / 1000);
        cache.put(key, JSON.stringify(payloadObj), ttlSec);

        Logger.log(`[${operation}] Key lock ACQUIRED "${lockKey}" lockId=${myId}`);
        return {
          scope: 'key',
          lockKey: key,
          userKey: lockKey,
          lockId: myId,
          operation,
          priority,
          lastHeartbeat: payloadObj.lastHeartbeat,
          expiry: payloadObj.expiry
        };
      }
    } catch (e) {
      Logger.log(`[${operation}] Guard/key lock internal error: ${e}`);
    } finally {
      try { guard.releaseLock(); } catch (_) { }
    }

    const cfg = BACKOFF_CFG[priority] || BACKOFF_CFG[3];
    const scaled = Math.min(cfg.base * Math.pow(1.5, attempt - 1), cfg.max);
    const jitter = Math.floor(Math.random() * 200);
    Utilities.sleep(scaled + jitter);
  }

  Logger.log(`[${operation}] FAILED to acquire key lock "${lockKey}" after ${Date.now() - start}ms`);
  return null;
}

/** Release either a sheet or key lock (polymorphic). */
function releaseDistributedLock(lockObj, operation = 'releaseDistributedLock') {
  if (!lockObj) return;
  const cache = CacheService.getScriptCache();
  const guard = LockService.getScriptLock();
  try {
    guard.waitLock(GUARD_WAIT_MS);
    const raw = cache.get(lockObj.lockKey);
    if (!raw) return;

    try {
      const info = JSON.parse(raw);
      if (info.lockId === lockObj.lockId) {
        cache.remove(lockObj.lockKey);
        Logger.log(`[${operation}] Released ${lockObj.scope || 'sheet'} lock "${lockObj.lockKey}".`);
      } else {
        Logger.log(`[${operation}] Ownership mismatch; not releasing.`);
      }
    } catch (e) {
      cache.remove(lockObj.lockKey);
      Logger.log(`[${operation}] Malformed data; force-removed "${lockObj.lockKey}".`);
    }
  } catch (e) {
    Logger.log(`[${operation}] Error acquiring guard to release: ${e}`);
  } finally {
    try { guard.releaseLock(); } catch (_) { }
  }
}

function updateLockHeartbeat(lockObj, operation = 'updateLockHeartbeat') {
  if (!lockObj || !lockObj.lockKey || !lockObj.lockId) {
    Logger.log(`[${operation}] Invalid lock object; cannot heartbeat.`);
    return false;
  }

  const cache = CacheService.getScriptCache();
  const guard = LockService.getScriptLock();
  try {
    // Serialize the read-modify-write to the cache entry
    guard.waitLock(GUARD_WAIT_MS);

    const raw = cache.get(lockObj.lockKey);
    if (!raw) {
      Logger.log(`[${operation}] Lock "${lockObj.lockKey}" not found in cache (maybe expired).`);
      return false;
    }

    let info;
    try {
      info = JSON.parse(raw);
    } catch (e) {
      Logger.log(`[${operation}] Malformed lock data for "${lockObj.lockKey}": ${e}`);
      return false;
    }

    // Ownership check
    if (info.lockId !== lockObj.lockId) {
      Logger.log(`[${operation}] Ownership mismatch for "${lockObj.lockKey}". Held by another lockId.`);
      return false;
    }

    // Refresh heartbeat + expiry
    const now = Date.now();
    info.lastHeartbeat = now;
    info.expiry = now + LOCK_TIMEOUT_MS;

    // (Optional) keep descriptive fields fresh if caller passed them
    if (lockObj.operation) info.operation = lockObj.operation;
    if (typeof lockObj.priority === 'number') info.priority = lockObj.priority;

    // Re-put with renewed TTL (lease + cushion)
    const ttlSec = Math.ceil((LOCK_TIMEOUT_MS + TTL_CUSHION_MS) / 1000);
    cache.put(lockObj.lockKey, JSON.stringify(info), ttlSec);

    // Reflect on the local object too (handy for callers)
    lockObj.lastHeartbeat = info.lastHeartbeat;
    lockObj.expiry = info.expiry;

    // Debug log (keep it lightweight)
    // Logger.log(`[${operation}] Heartbeat OK for "${lockObj.lockKey}". Next expiry in ${LOCK_TIMEOUT_MS}ms.`);

    return true;

  } catch (e) {
    Logger.log(`[${operation}] Error while heartbeating "${lockObj.lockKey}": ${e}`);
    return false;
  } finally {
    try { guard.releaseLock(); } catch (_) { }
  }
}

/** Heartbeat wrapper usable for sheet or key locks. */
function heartbeatLock(lockObj) {
  return updateLockHeartbeat(lockObj);
}

/**
 * Run a critical section with a key-based distributed lock.
 * If your section could be long, call the provided `beat()` to extend the lease as needed.
 */
function withKeyLock(lockKey, operation, fn, priority = 2, maxWaitMs = 60000) {
  const lock = acquireKeyLock(lockKey, operation, priority, maxWaitMs);
  if (!lock) {
    throw new Error(`[${operation}] Could not acquire key lock "${lockKey}" within ${maxWaitMs}ms`);
  }
  try {
    return fn(lock, () => heartbeatLock(lock));
  } finally {
    releaseDistributedLock(lock, operation);
  }
}

/**************************************
 * wrapper_utils.js (row-lock shim)
 **************************************/

/** Build a stable user-visible key for a row lock (the `makeKeyLockKey` will sanitize it). */
function makeRowUserKey(sheetName, rowIndex) {
  return `row:${String(sheetName)}:${Number(rowIndex)}`;
}

/**
 * Acquire a row-level lock (wrapper over acquireKeyLock).
 * priority: lower number = higher priority. e.g. edits=1, interval=2.
 */
function acquireRowLock(sheetName, rowIndex, operation, priority = 2, maxWaitMs = 60000) {
  const userKey = makeRowUserKey(sheetName, rowIndex);
  return acquireKeyLock(userKey, `${operation}@${sheetName}#${rowIndex}`, priority, maxWaitMs);
}

/** Release a row-level lock (polymorphic release already works). */
function releaseRowLock(lockObj, operation = 'releaseRowLock') {
  return releaseDistributedLock(lockObj, operation);
}

/** Heartbeat/extend row lock TTL. */
function heartbeatRowLock(lockObj) {
  return heartbeatLock(lockObj); // your existing update/heartbeat logic
}

/**
 * Run a critical section while holding a row-level lock.
 * `fn(lock, beat)` receives:
 *  - lock: the lock object
 *  - beat(): call periodically if your work can exceed LOCK_TIMEOUT_MS
 */
function withRowLock(sheetName, rowIndex, operation, fn, priority = 2, maxWaitMs = 60000) {
  const lock = acquireRowLock(sheetName, rowIndex, operation, priority, maxWaitMs);
  if (!lock) {
    throw new Error(`[${operation}] Could not acquire row lock for ${sheetName}#${rowIndex} within ${maxWaitMs}ms`);
  }
  try {
    return fn(lock, () => heartbeatRowLock(lock));
  } finally {
    releaseRowLock(lock, operation);
  }
}