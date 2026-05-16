// ============================================================
//  MWAAF – DEVIATION AUTHORIZATION SYSTEM
//  Google Apps Script Backend  |  Code.gs
//  
//  SHEET NAMES (must match exactly):
//    "Deviations"     — one row per deviation
//    "Approvals"      — one row per approval decision
//    "Approvers"      — configured approvers + roles
//    "DistLists"      — email distribution lists
//    "Config"         — app-wide settings
//    "ApprovalTokens" — secure tokens for email links
// ============================================================

// ── CONFIGURATION ──────────────────────────────────────────
// NO manual URL needed — the script detects its own deployed URL automatically.
// Just deploy as Web App and it works. No need to update this file after each deploy.

// Lazy getters prevent crashes when Apps Script calls doGet/doPost
// without an active spreadsheet context (e.g. when run from editor).
function getSheetId() {
  const stored = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (stored) return stored;
  try {
    const id = SpreadsheetApp.getActiveSpreadsheet().getId();
    PropertiesService.getScriptProperties().setProperty('SHEET_ID', id);
    return id;
  } catch(e) {
    throw new Error('SHEET_ID not set. Run setupSheets() first from the spreadsheet editor.');
  }
}

function getWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
}

// Sheet name constants
// ── EMAIL SENDER CONFIGURATION ──────────────────────────────────
// To change the "From" address, you must FIRST add it as an alias in Gmail Settings:
//   Settings > Accounts > "Send mail as" > "Add another email address"
// Once verified, set FROM_EMAIL below. If empty, emails send from the script owner's address.
const FROM_EMAIL = '';   // ← LEAVE EMPTY until DeviationAuth@mwaaf.com is added as a Gmail alias.
                         //    To enable: 1) Add alias in Gmail Settings > Accounts > Send mail as
                         //              2) Verify it via the email Google sends
                         //              3) Then set this to 'DeviationAuth@mwaaf.com'
const FROM_NAME  = 'MWAAF Deviation System';

const SH = {
  DEVIATIONS:      'Deviations',
  APPROVALS:       'Approvals',
  APPROVERS:       'Approvers',
  DIST_LISTS:      'DistLists',
  CONFIG:          'Config',
  TOKENS:          'ApprovalTokens',
  ROLE_CONFIG:     'RoleConfig',
  REASON_OPTIONS:  'ReasonOptions',
  PART_NUMBERS:    'PartNumbers',
  WORK_CENTERS:    'WorkCenters',
};

// ── ENTRY POINTS ───────────────────────────────────────────

/**
 * GET handler — serves the approver view or the main app.
 * URL params:
 *   ?token=XXX          → approver's personalized view
 *   ?page=app           → main app (HTML)
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  // Approver view via secure token link (sent in approval emails)
  if (params.token) {
    return serveApproverView(params.token);
  }

  // Read-only deviation view (sent in approval-complete emails so approvers can verify)
  // Format: ?view=DEV-1004
  if (params.view) {
    return serveReadOnlyDeviationView(params.view);
  }

  // Default: status page (useful to verify the Web App is live)
  return HtmlService.createHtmlOutput(
    '<html><body style="font-family:Calibri,sans-serif;padding:30px;max-width:500px">' +
    '<h2 style="color:#7A9A3A">&#x2705; MWAAF Backend is running</h2>' +
    '<p>Web App is deployed and responding correctly.</p>' +
    '<p>Open <strong>deviation-app-sheets.html</strong> in your browser to use the app.</p>' +
    '<hr><p style="color:#999;font-size:12px">Checked at: ' + new Date().toISOString() + '</p>' +
    '</body></html>'
  ).setTitle('MWAAF Backend');
}

/**
 * POST handler — all API calls from the front-end.
 * Body: { action: 'string', payload: {...} }
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'No POST data received. This endpoint requires HTTP POST requests.' });
    }
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};

    const handlers = {
      // Deviations CRUD
      'getDeviations':       () => getDeviations(payload),
      'getDeviationsLite':   () => getDeviationsLite(payload),
      'getDeviationPhotos':  () => getDeviationPhotos(payload),
      'saveDeviation':       () => saveDeviation(payload),
      'deleteDeviation':     () => deleteDeviation(payload.id),

      // Approvals
      'submitApproval':      () => submitApproval(payload),

      // Auth
      'login':               () => loginApprover(payload),

      // Approvers management
      'getApprovers':        () => getApprovers(),
      'saveApprover':        () => saveApprover(payload),
      'saveApprovers':       () => saveApprovers(payload),
      'removeApprover':      () => removeApprover(payload.id),

      // Config
      'getConfig':           () => getFullConfig(),
      'saveConfig':          () => saveConfig(payload),
      'saveRoleConfig':      () => saveRoleConfig(payload),
      'saveReasonOptions':   () => saveReasonOptions(payload),
      'saveDistLists':       () => saveDistLists(payload),

      // Part Numbers & Work Centers (catalogs)
      'getPartNumbers':      () => getPartNumbers(),
      'savePartNumbers':     () => savePartNumbers(payload),
      'upsertPartNumber':    () => upsertPartNumber(payload),
      'getWorkCenters':      () => getWorkCenters(),
      'saveWorkCenters':     () => saveWorkCenters(payload),

      // Notifications
      'sendNotification':    () => sendNotification(payload),
      'getEmailQuota':       () => getEmailQuota(),

      // App access gate (HTML splash password)
      'verifyAppAccess':     () => verifyAppAccess(payload),
      'getAppLockStatus':    () => getAppLockStatus(),

      // Settings tab gate (in-app password)
      'verifySettingsPassword': () => verifySettingsPassword(payload),
      'getSettingsLockStatus':  () => getSettingsLockStatus(),
    };

    if (!handlers[action]) {
      return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }

    const result = handlers[action]();
    return jsonResponse({ ok: true, data: result });

  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── HELPERS ────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, createIfMissing) {
  const ss = SpreadsheetApp.openById(getSheetId());
  let sheet = ss.getSheetByName(name);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/** Reads all rows of a sheet as array of objects using row 1 as headers.
 *  Converts Date objects to ISO strings so JS clients receive plain strings,
 *  not Date objects that lack .replace() and similar string methods.
 */
function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName, false);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  // Get the spreadsheet's timezone so dates are interpreted in the user's local context
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/New_York';
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      // Convert Date objects → ISO string so clients always get plain strings
      if (v instanceof Date) {
        if (v.getTime() === 0) {
          obj[h] = '';
        } else {
          // For date-only fields (date, startDate, endDate), output YYYY-MM-DD using the SPREADSHEET timezone
          // (Sheets stores 2026-04-25 as midnight in the sheet's tz; toISOString() shifts to UTC and may drop a day)
          // For datetime fields (submittedAt), keep the full ISO string
          if (h === 'date' || h === 'startDate' || h === 'endDate') {
            obj[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
          } else {
            obj[h] = v.toISOString();
          }
        }
      } else {
        obj[h] = v;
      }
    });
    return obj;
  });
}

/** Defensive string conversion: handles null, undefined, Date, number, etc. */
function toStr(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (v instanceof Date) return v.toISOString();
  return String(v);
}


/** Format YYYY-MM-DD as MM/DD/YY for emails and approver view */
function fmtUS(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/New_York';
    return Utilities.formatDate(v, tz, 'MM/dd/yy');
  }
  if (typeof v === 'string') {
    const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[2] + '/' + m[3] + '/' + m[1].slice(-2);
  }
  return String(v);
}

/** Defensive newline-to-<br> replacement that won't crash on non-strings */
function nlToBr(v) {
  return toStr(v).replace(/\n/g, '<br>');
}

/** Writes an array of objects to a sheet, replacing all content. */
function objectsToSheet(sheetName, objects, headers) {
  const sheet = getSheet(sheetName, true);
  sheet.clearContents();
  if (!objects.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const rows = [headers, ...objects.map(o => headers.map(h => o[h] !== undefined ? o[h] : ''))];
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  // Style header row
  styleHeader(sheet, headers.length);
}

function styleHeader(sheet, numCols) {
  const hdrRange = sheet.getRange(1, 1, 1, numCols);
  hdrRange.setBackground('#7A9A3A')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setFontSize(11);
}

function generateId(prefix) {
  return prefix.toLowerCase() + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g, '');
}

// ── DEVIATIONS ─────────────────────────────────────────────

const DEV_HEADERS = [
  'id','devNum','mainPartNum','date','shift','initiator','workCenter',
  'custApproval','startDate','endDate','description',
  'parts','reasons','reasonLabel','tags','fourm','riskFactor',
  'actionPlan','comments','owner','otherReason','status',
  'submittedAt','revision','photos','selectedApprovers'
];

const APPR_HEADERS = [
  'id','deviationId','approverId','approverName','approverRole',
  'decision','date','comments'
];

function getDeviations(filter) {
  let devs = sheetToObjects(SH.DEVIATIONS);
  const approvals = sheetToObjects(SH.APPROVALS);

  // Parse JSON fields
  devs = devs.map(d => {
    ['parts','reasons','tags','fourm','photos','selectedApprovers'].forEach(f => {
      if (typeof d[f] === 'string' && d[f]) {
        try { d[f] = JSON.parse(d[f]); } catch(e) { d[f] = []; }
      }
      if (!d[f]) d[f] = [];
    });
    // Attach approvals
    d.approvals = approvals
      .filter(a => a.deviationId === d.id)
      .map(a => ({
        approverId: a.approverId,
        decision:   a.decision,
        date:       a.date,
        comments:   a.comments,
      }));
    return d;
  });

  // Optional filters
  if (filter) {
    if (filter.status) devs = devs.filter(d => d.status === filter.status);
    if (filter.workCenter) devs = devs.filter(d => String(d.workCenter).toLowerCase().includes(filter.workCenter.toLowerCase()));
    if (filter.approverId) {
      // Only deviations still pending for this approver
      devs = devs.filter(d =>
        (d.status === 'pending' || d.status === 'partial') &&
        !d.approvals.find(a => String(a.approverId || '').toLowerCase() === String(filter.approverId || '').toLowerCase())
      );
    }
  }

  return devs;
}

/**
 * "Lite" deviations: same as getDeviations but strips photo dataURLs.
 * Each photo becomes { hasPhoto: true } with no dataUrl, so the dashboard
 * still knows there are photos (for the count badge) but doesn't pay the
 * cost of transferring potentially MB of base64 per row.
 *
 * Use this for the dashboard. Use getDeviationPhotos(id) when the user
 * actually opens a specific deviation.
 *
 * Performance: 500 deviations with photos: ~30-90s with full → ~3-8s with lite.
 */
function getDeviationsLite(filter) {
  const devs = getDeviations(filter);
  return devs.map(d => {
    const photoCount = Array.isArray(d.photos) ? d.photos.length : 0;
    // Replace photos array with a placeholder array of the same length but no dataURLs.
    // The dashboard only needs to know "has photos" / count for the UI badge.
    d.photos = [];
    d._photoCount = photoCount;
    return d;
  });
}

/**
 * Returns just the photos for a single deviation, by id.
 * Called lazily when the user opens a deviation detail view.
 *
 * Returns: { id, photos: [{dataUrl, ...}] } or null if not found.
 */
function getDeviationPhotos(payload) {
  const id = String((payload && payload.id) || '').trim();
  if (!id) return null;
  const sheet = getSheet(SH.DEVIATIONS, false);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data.length) return null;
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const photosCol = headers.indexOf('photos');
  if (idCol < 0 || photosCol < 0) return null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      let photos = data[i][photosCol];
      if (typeof photos === 'string' && photos) {
        try { photos = JSON.parse(photos); } catch(e) { photos = []; }
      }
      if (!Array.isArray(photos)) photos = [];
      return { id: id, photos: photos };
    }
  }
  return null;
}

function saveDeviation(dev) {
  // ── INPUT VALIDATION (defense-in-depth) ──
  // Reject pathologically-large or malformed payloads before they hit the sheet.
  if (!dev || typeof dev !== 'object') {
    throw new Error('Invalid payload: deviation must be an object.');
  }
  // Field length caps to prevent abuse / sheet bloat. These are very generous —
  // legitimate deviations stay well under these limits.
  const MAX_TEXT = 10000;     // 10K chars per text field
  const MAX_LONG_TEXT = 30000; // 30K for description/actionPlan/comments
  const MAX_PARTS = 200;      // max affected parts entries
  const MAX_PHOTOS = 20;      // max photos
  const MAX_PHOTO_BYTES = 5 * 1024 * 1024; // 5MB per photo (data URL approximate)
  const _checkLen = (v, max, name) => {
    if (v !== null && v !== undefined && String(v).length > max) {
      throw new Error('Field "' + name + '" exceeds maximum length (' + max + ' chars).');
    }
  };
  ['devNum','mainPartNum','workCenter','initiator','reasonLabel','riskFactor','custApproval','owner']
    .forEach(f => _checkLen(dev[f], MAX_TEXT, f));
  ['description','actionPlan','comments','otherReason']
    .forEach(f => _checkLen(dev[f], MAX_LONG_TEXT, f));
  if (Array.isArray(dev.parts) && dev.parts.length > MAX_PARTS) {
    throw new Error('Too many affected parts (max ' + MAX_PARTS + ').');
  }
  if (Array.isArray(dev.photos)) {
    if (dev.photos.length > MAX_PHOTOS) {
      throw new Error('Too many photos (max ' + MAX_PHOTOS + ').');
    }
    // Validate each photo dataUrl
    for (let i = 0; i < dev.photos.length; i++) {
      const p = dev.photos[i];
      if (!p || typeof p !== 'object') continue;
      if (p.dataUrl) {
        const url = String(p.dataUrl);
        if (!/^data:image\/(png|jpe?g|gif|webp);base64,/i.test(url)) {
          throw new Error('Photo ' + (i+1) + ': only image data URLs (PNG/JPEG/GIF/WebP) are allowed.');
        }
        if (url.length > MAX_PHOTO_BYTES) {
          throw new Error('Photo ' + (i+1) + ' exceeds size limit (5MB).');
        }
      }
    }
  }
  // Status whitelist
  const validStatuses = ['draft','pending','partial','approved','rejected','expired',''];
  if (dev.status && validStatuses.indexOf(dev.status) === -1) {
    throw new Error('Invalid status value: ' + dev.status);
  }
  // Risk whitelist
  const validRisks = ['Low','Med','High',''];
  if (dev.riskFactor && validRisks.indexOf(dev.riskFactor) === -1) {
    throw new Error('Invalid risk factor: ' + dev.riskFactor);
  }
  // ── END VALIDATION ──

  const sheet = getSheet(SH.DEVIATIONS, true);
  const data = sheet.getDataRange().getValues();

  // Ensure headers exist
  if (!data.length || data[0][0] !== 'id') {
    sheet.clearContents();
    sheet.appendRow(DEV_HEADERS);
    styleHeader(sheet, DEV_HEADERS.length);
  }

  // Stringify JSON fields
  const row = { ...dev };
  ['parts','reasons','tags','fourm','photos','selectedApprovers'].forEach(f => {
    if (Array.isArray(row[f])) row[f] = JSON.stringify(row[f]);
  });
  delete row.approvals; // stored separately

  // Find existing row FIRST (before generating any IDs)
  // This is critical: the frontend always sends a temporary id like 'dev_<timestamp>'
  // for new deviations, so we cannot rely on `!row.id` to detect new ones.
  // Instead, check if the id is already in the sheet — if not, treat as new.
  const currentData = sheet.getDataRange().getValues();
  const headers = currentData[0];
  const idCol = headers.indexOf('id');
  const existingRowIdx = currentData.slice(1).findIndex(r => r[idCol] === row.id);
  const isNew = existingRowIdx < 0;

  // Generate canonical ID and consecutive devNum for new deviations
  if (isNew) {
    if (!row.id) row.id = generateId('dev');
    // Always assign a fresh consecutive devNum from the persistent counter,
    // regardless of what the frontend sent (which is just a placeholder).
    row.devNum = generateNextDevNum();
    row.submittedAt = row.submittedAt || new Date().toISOString();
    row.revision = row.revision || 1;
    Logger.log('saveDeviation: NEW deviation, assigned devNum=' + row.devNum + ', id=' + row.id);
  } else {
    Logger.log('saveDeviation: UPDATE existing deviation id=' + row.id + ', keeping devNum=' + row.devNum);
  }

  const rowValues = DEV_HEADERS.map(h => row[h] !== undefined ? row[h] : '');

  if (!isNew) {
    // Update existing row (slice(1) offset + 1-based + header row = +2)
    sheet.getRange(existingRowIdx + 2, 1, 1, DEV_HEADERS.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  // ── EDIT FLOW: re-trigger approval workflow ──
  // When an existing deviation is edited (and not just being saved as a draft),
  // wipe all previous Approvals so the approvers must sign again on the updated
  // content. Also reset status to 'pending' so the dashboard reflects that
  // approvals are required again, and bump the revision number so approvers
  // can see this is a new cycle.
  let approvalsCleared = 0;
  let newRevision = null;
  if (!isNew && row.status !== 'draft') {
    try {
      approvalsCleared = _clearApprovalsForDeviation(row.id);
      // Reset status on the just-updated row to 'pending' (overwrites whatever
      // status came in from the frontend).
      const statusCol = DEV_HEADERS.indexOf('status');
      if (statusCol >= 0) {
        sheet.getRange(existingRowIdx + 2, statusCol + 1).setValue('pending');
      }
      // Bump revision by 1 (defaults to 1 if missing)
      const revCol = DEV_HEADERS.indexOf('revision');
      if (revCol >= 0) {
        const currentRev = parseInt(rowValues[revCol], 10);
        newRevision = (isNaN(currentRev) ? 1 : currentRev) + 1;
        sheet.getRange(existingRowIdx + 2, revCol + 1).setValue(newRevision);
      }
      Logger.log('saveDeviation: edit re-triggered approval workflow. Cleared ' + approvalsCleared + ' previous approvals; status reset to pending; revision bumped to ' + newRevision + '.');
    } catch(e) {
      Logger.log('saveDeviation: failed to clear approvals on edit: ' + e.message);
    }
  }

  return { id: row.id, devNum: row.devNum, isNew: isNew, approvalsCleared: approvalsCleared, newRevision: newRevision };
}

/**
 * Removes all Approvals rows for a given deviationId.
 * Returns the number of rows deleted. Used when a deviation is edited and
 * approvers must re-sign on the updated content.
 */
function _clearApprovalsForDeviation(deviationId) {
  const sheet = getSheet(SH.APPROVALS, false);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  const headers = data[0];
  const devCol = headers.indexOf('deviationId');
  if (devCol < 0) return 0;
  let deleted = 0;
  // Walk bottom-up so row indices stay valid as we delete.
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][devCol] === deviationId) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}

function deleteDeviation(id) {
  const sheet = getSheet(SH.DEVIATIONS, false);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function generateNextDevNum() {
  // Use script-level lock to prevent two concurrent submissions from getting the same folio
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // wait up to 10 seconds
  } catch(e) {
    Logger.log('generateNextDevNum: could not acquire lock: ' + e.message);
  }
  try {
    const configSheet = getSheet(SH.CONFIG, true);
    const data = configSheet.getDataRange().getValues();
    let nextNum = 1001;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'nextDevNum') {
        nextNum = parseInt(data[i][1]) || 1001;
        configSheet.getRange(i + 1, 2).setValue(nextNum + 1);
        SpreadsheetApp.flush();
        return 'DEV-' + nextNum;
      }
    }
    // Not found — init
    configSheet.appendRow(['nextDevNum', nextNum + 1]);
    SpreadsheetApp.flush();
    return 'DEV-' + nextNum;
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ── APPROVALS ──────────────────────────────────────────────

function submitApproval(payload) {
  // payload: { deviationId, approverId, decision, comments, approverToken }
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid payload.');
  }
  // Validate decision (whitelist)
  const validDecisions = ['approved','rejected'];
  if (validDecisions.indexOf(payload.decision) === -1) {
    throw new Error('Invalid decision value.');
  }
  // Cap comments length
  if (payload.comments && String(payload.comments).length > 5000) {
    throw new Error('Comments exceed maximum length (5000 chars).');
  }
  // Basic format checks for IDs
  if (!payload.deviationId || !/^[a-zA-Z0-9_\-]+$/.test(String(payload.deviationId))) {
    throw new Error('Invalid deviation ID format.');
  }
  if (!payload.approverId || String(payload.approverId).length > 200) {
    throw new Error('Invalid approver ID.');
  }

  // CRITICAL: normalize approverId to lowercase so id matches the Approvers sheet
  payload.approverId = String(payload.approverId || '').toLowerCase().trim();

  // Validate token
  if (payload.approverToken) {
    const tokenData = validateToken(payload.approverToken);
    if (!tokenData) throw new Error('Invalid or expired approval token.');
    if (String(tokenData.approverId).toLowerCase() !== payload.approverId) throw new Error('Token mismatch.');
  }

  const sheet = getSheet(SH.APPROVALS, true);
  const data = sheet.getDataRange().getValues();
  if (!data.length || data[0][0] !== 'id') {
    sheet.clearContents();
    sheet.appendRow(APPR_HEADERS);
    styleHeader(sheet, APPR_HEADERS.length);
  }

  // Check not already approved
  const approvers = getApprovers();
  const approver = approvers.find(a => String(a.id).toLowerCase() === payload.approverId);
  if (!approver) throw new Error('Approver not found.');

  // Check duplicate (case-insensitive)
  const existing = sheetToObjects(SH.APPROVALS);
  if (existing.find(a => a.deviationId === payload.deviationId && String(a.approverId).toLowerCase() === payload.approverId)) {
    throw new Error('You have already submitted a decision for this deviation.');
  }

  // Get deviation status BEFORE recording this decision (so we can detect informational cases)
  const allDevs = sheetToObjects(SH.DEVIATIONS);
  const devBefore = allDevs.find(d => d.id === payload.deviationId) || {};
  const statusBefore = devBefore.status || 'pending';

  // Determine this approver's role status (required/optional/ruleout) in this deviation
  let selectedApprovers = [];
  if (devBefore.selectedApprovers) {
    try {
      selectedApprovers = typeof devBefore.selectedApprovers === 'string'
        ? JSON.parse(devBefore.selectedApprovers)
        : devBefore.selectedApprovers;
    } catch(e) { selectedApprovers = []; }
  }
  selectedApprovers = (selectedApprovers || []).filter(s => s && s.checked !== false);
  if (!selectedApprovers.length) {
    selectedApprovers = approvers.map(a => ({
      id: a.id,
      checked: true,
      status: a.defaultStatus || (a.required ? 'required' : 'optional')
    }));
  }
  const myEntry = selectedApprovers.find(s => String(s.id || '').toLowerCase() === payload.approverId);
  const myStatus = myEntry ? (myEntry.status || 'optional') : 'optional';

  const apprRecord = {
    id:            generateId('appr'),
    deviationId:   payload.deviationId,
    approverId:    payload.approverId,
    approverName:  approver.name,
    approverRole:  approver.role,
    decision:      payload.decision,
    date:          new Date().toISOString(),
    comments:      payload.comments || '',
  };
  sheet.appendRow(APPR_HEADERS.map(h => apprRecord[h] || ''));

  // Update deviation status (uses new priority rules: ruleout-approved wins over rejection,
  // optional rejection doesn't block, etc.)
  updateDeviationStatus(payload.deviationId);

  // Read status AFTER to compose feedback
  const allDevsAfter = sheetToObjects(SH.DEVIATIONS);
  const devAfter = allDevsAfter.find(d => d.id === payload.deviationId) || {};
  const statusAfter = devAfter.status || 'pending';

  // Build a human-readable info message for the approver
  let info = '';
  if (payload.decision === 'rejected' && myStatus === 'optional' && (statusBefore === 'approved' || statusAfter === 'approved')) {
    info = 'This deviation was already approved by required approvers. Your rejection has been recorded as a comment but does not change the deviation status.';
  } else if (payload.decision === 'approved' && myStatus === 'ruleout' && statusBefore === 'rejected') {
    info = 'Your rule-out approval has overridden a previous rejection. The deviation is now APPROVED.';
  } else if (payload.decision === 'rejected' && myStatus === 'optional') {
    info = 'Your rejection was recorded for the audit log, but optional rejections do not block approval — the deviation continues to need required signatures.';
  }

  // Mark token as used
  if (payload.approverToken) {
    markTokenUsed(payload.approverToken);
  }

  return {
    record: apprRecord,
    statusBefore: statusBefore,
    statusAfter: statusAfter,
    myStatus: myStatus,
    info: info
  };
}

/**
 * Called directly from ApproverView.html via google.script.run
 * No fetch/CORS needed — runs server-side in the same script context.
 */
function submitApprovalFromView(deviationId, approverId, decision, comments, token) {
  return submitApproval({
    deviationId:   deviationId,
    approverId:    approverId,
    decision:      decision,
    comments:      comments,
    approverToken: token,
  });
}


/* Helper: read a single deviation row from the sheet by id, without
   loading and parsing all 500+ rows. Returns the raw object with parsed
   selectedApprovers if present. Used by updateDeviationStatus to avoid
   the O(N) scan for status updates. */
function _getOneDeviationRaw(deviationId) {
  const sheet = getSheet(SH.DEVIATIONS, false);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data.length) return null;
  const headers = data[0].map(h => String(h));
  const idCol = headers.indexOf('id');
  if (idCol < 0) return null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === deviationId) {
      const row = {};
      for (let c = 0; c < headers.length; c++) {
        row[headers[c]] = data[i][c];
      }
      // Parse selectedApprovers (the only JSON field updateDeviationStatus needs)
      if (typeof row.selectedApprovers === 'string' && row.selectedApprovers) {
        try { row.selectedApprovers = JSON.parse(row.selectedApprovers); } catch(e) { row.selectedApprovers = []; }
      }
      return row;
    }
  }
  return null;
}

function updateDeviationStatus(deviationId) {
  // Only read approvals for THIS deviation, not all
  const approvalSheet = getSheet(SH.APPROVALS, false);
  const approvals = [];
  if (approvalSheet) {
    const aData = approvalSheet.getDataRange().getValues();
    if (aData.length > 1) {
      const aHeaders = aData[0].map(h => String(h));
      const aDevIdCol = aHeaders.indexOf('deviationId');
      const aDecisionCol = aHeaders.indexOf('decision');
      const aApproverIdCol = aHeaders.indexOf('approverId');
      if (aDevIdCol >= 0 && aDecisionCol >= 0 && aApproverIdCol >= 0) {
        for (let i = 1; i < aData.length; i++) {
          if (String(aData[i][aDevIdCol]) === deviationId) {
            approvals.push({
              deviationId: aData[i][aDevIdCol],
              approverId: aData[i][aApproverIdCol],
              decision: aData[i][aDecisionCol],
            });
          }
        }
      }
    }
  }

  // Get only the single deviation row (instead of loading all 500+)
  const dev = _getOneDeviationRaw(deviationId);
  let selectedApprovers = [];
  if (dev && dev.selectedApprovers) {
    selectedApprovers = Array.isArray(dev.selectedApprovers) ? dev.selectedApprovers : [];
  }

  // Normalize: only consider checked entries
  selectedApprovers = (selectedApprovers || []).filter(s => s && s.checked !== false);

  // Fallback: if no per-deviation selection, use Approvers sheet defaults
  if (!selectedApprovers.length) {
    const allApprovers = getApprovers();
    selectedApprovers = allApprovers.map(a => ({
      id: a.id,
      checked: true,
      status: a.defaultStatus || (a.required ? 'required' : 'optional')
    }));
  }

  // Build lookups
  const statusById = {};
  selectedApprovers.forEach(s => {
    statusById[String(s.id || '').toLowerCase()] = s.status || 'optional';
  });

  const approvedIds = approvals.filter(a => a.decision === 'approved').map(a => String(a.approverId || '').toLowerCase());
  const rejectedIds = approvals.filter(a => a.decision === 'rejected').map(a => String(a.approverId || '').toLowerCase());

  const ruleouts = selectedApprovers.filter(s => s.status === 'ruleout');
  const required = selectedApprovers.filter(s => s.status === 'required');

  // ── PRIORITY ORDER (highest → lowest) ──
  // The new rules:
  //   1. Ruleout APPROVED → approved (overrides any prior rejection)
  //   2. All required APPROVED → approved (optional rejections after this don't matter)
  //   3. Any REQUIRED rejected → rejected (a required saying "no" blocks the deviation)
  //   4. Some required approved → partial
  //   5. Otherwise → pending
  //
  // Note: rejections from OPTIONAL approvers are recorded in the audit log but do
  // NOT change the deviation status. Rejections from required DO block.

  // RULE 1: Any ruleout approved → instant approval (wins over rejections)
  const ruleOutApproved = ruleouts.some(s => approvedIds.includes(String(s.id || '').toLowerCase()));
  if (ruleOutApproved) {
    _setDeviationStatus(deviationId, 'approved');
    _sendApprovalEmail(deviationId);
    return;
  }

  // RULE 2: All required signed → approved
  if (required.length > 0) {
    const allRequiredApproved = required.every(s => approvedIds.includes(String(s.id || '').toLowerCase()));
    if (allRequiredApproved) {
      _setDeviationStatus(deviationId, 'approved');
      _sendApprovalEmail(deviationId);
      return;
    }
  }

  // RULE 3: Any REQUIRED rejection → rejected
  // Optional rejections are noted but not status-changing.
  const anyRequiredRejected = rejectedIds.some(rid => statusById[rid] === 'required');
  if (anyRequiredRejected) {
    _setDeviationStatus(deviationId, 'rejected');
    _sendRejectedEmail(deviationId);
    return;
  }

  // RULE 4: Some required approved (but not all) → partial
  if (required.length > 0) {
    const someRequiredApproved = required.some(s => approvedIds.includes(String(s.id || '').toLowerCase()));
    if (someRequiredApproved) {
      _setDeviationStatus(deviationId, 'partial');
      return;
    }
  }

  // RULE 5: No required, no ruleouts — fallback: any approval = approved
  if (required.length === 0 && ruleouts.length === 0 && approvedIds.length > 0) {
    _setDeviationStatus(deviationId, 'approved');
    _sendApprovalEmail(deviationId);
    return;
  }

  _setDeviationStatus(deviationId, 'pending');
}

function _setDeviationStatus(deviationId, newStatus) {
  const sheet = getSheet(SH.DEVIATIONS, false);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const statusCol = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === deviationId) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      return;
    }
  }
}

function _sendApprovalEmail(deviationId) {
  try {
    const devs = getDeviations({});
    const updatedDev = devs.find(d => d.id === deviationId);
    if (updatedDev) sendNotification({ dev: updatedDev, type: 'approval' });
  } catch(e) { Logger.log('approval email failed: ' + e); }
}

function _sendRejectedEmail(deviationId) {
  try {
    const devs = getDeviations({});
    const updatedDev = devs.find(d => d.id === deviationId);
    if (updatedDev) sendNotification({ dev: updatedDev, type: 'rejected' });
  } catch(e) { Logger.log('rejected email failed: ' + e); }
}

/**
 * Returns the remaining MailApp quota for today and the assumed daily limit.
 * Consumer Gmail accounts get 100/day; Workspace gets 1500/day. We don't
 * know which one is hosting the script, so we infer the limit from the
 * current remaining value (if remaining > 100, it must be Workspace).
 */
function getEmailQuota() {
  try {
    const remaining = MailApp.getRemainingDailyQuota();
    // Heuristic: consumer = 100/day, Workspace = 1500/day.
    // If remaining is over 100, we're definitely on Workspace.
    const limit = remaining > 100 ? 1500 : 100;
    return { remaining: remaining, limit: limit, used: Math.max(0, limit - remaining) };
  } catch(e) {
    Logger.log('getEmailQuota failed: ' + e.message);
    return { remaining: null, limit: 100, used: null, error: e.message };
  }
}

// ── AUTH / LOGIN ───────────────────────────────────────────

function loginApprover(payload) {
  // payload: { email, password }
  // NOTE: this is a legacy login flow; current production uses token-based access.
  const approvers = getApprovers();
  const approver = approvers.find(a => a.email.toLowerCase() === payload.email.toLowerCase());
  const config = getConfig();
  // Use settingsPassword (same key used by the Settings tab gate).
  const expected = String(config.settingsPassword || '').trim();
  if (!expected) {
    // No password configured — reject (legacy flow shouldn't be allowed without a configured password)
    throw new Error('Login is not enabled. Use the personalized email link instead.');
  }
  if (!approver || payload.password !== expected) {
    throw new Error('Invalid credentials.');
  }
  return {
    id:    approver.id,
    name:  approver.name,
    email: approver.email,
    role:  approver.role,
  };
}

// ── TOKENS ─────────────────────────────────────────────────

/**
 * Creates a unique approval token for an approver.
 * Tokens are stored in ApprovalTokens sheet with expiry.
 */
function createApproverToken(approverId, approverEmail) {
  const sheet = getSheet(SH.TOKENS, true);
  const tokenHeaders = ['token','approverId','approverEmail','createdAt','expiresAt','used'];
  const data = sheet.getDataRange().getValues();
  if (!data.length || data[0][0] !== 'token') {
    sheet.clearContents();
    sheet.appendRow(tokenHeaders);
    styleHeader(sheet, tokenHeaders.length);
  }

  // Force ENTIRE sheet to PLAIN TEXT format permanently to prevent date auto-conversion
  sheet.getRange(1, 1, sheet.getMaxRows(), tokenHeaders.length).setNumberFormat('@');

  const token = generateToken();
  const now = new Date();
  const expires = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // 7 days
  
  // Sanitize approverId — if it looks like a date (corrupted), rebuild as 'apr_<lowercase>'
  let safeApproverId = String(approverId || '').toLowerCase().trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(safeApproverId) || /^\d+\/\d+\/\d{4}/.test(safeApproverId)) {
    Logger.log('createApproverToken: approverId was date-like (' + safeApproverId + '), needs cleaning by caller');
  }
  // PREFIX with apostrophe to force Sheets to treat as text literal
  // (apostrophe trick: leading single quote in a cell = text override)
  const idForSheet = "'" + safeApproverId;

  // Use appendRow with pre-formatted range (text mode set above)
  const lastRow = sheet.getLastRow() + 1;
  const range = sheet.getRange(lastRow, 1, 1, 6);
  range.setNumberFormat('@');
  
  // Use setValues with explicit text values
  range.setValues([[
    "'" + token,
    idForSheet,
    "'" + (approverEmail || ''),
    "'" + now.toISOString(),
    "'" + expires.toISOString(),
    "'false"
  ]]);

  return token;
}

function validateToken(token) {
  const sheet = getSheet(SH.TOKENS, false);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const tokenCol   = headers.indexOf('token');
  const aprCol     = headers.indexOf('approverId');
  const emailCol   = headers.indexOf('approverEmail');
  const expiryCol  = headers.indexOf('expiresAt');
  const usedCol    = headers.indexOf('used');

  // Helper: strip leading apostrophe text-prefix
  function stripPrefix(v){
    if (v === null || v === undefined) return '';
    let s = String(v);
    if (s.charAt(0) === "'") s = s.substring(1);
    return s;
  }

  // Normalize the incoming token (may also have apostrophe if compared against raw cell)
  const tokenLookup = String(token || '').trim();

  for (let i = 1; i < data.length; i++) {
    const cellToken = stripPrefix(data[i][tokenCol]).trim();
    if (cellToken === tokenLookup) {
      const cellUsed = stripPrefix(data[i][usedCol]).toLowerCase();
      if (cellUsed === 'true') {
        Logger.log('Token row ' + i + ': already used');
        return null;
      }
      
      // Parse expiry — handle string, Date, or stripped string
      const cellExpiry = stripPrefix(data[i][expiryCol]);
      let expires;
      if (data[i][expiryCol] instanceof Date) {
        expires = data[i][expiryCol];
      } else {
        expires = new Date(cellExpiry);
      }
      if (isNaN(expires.getTime())) {
        Logger.log('Token row ' + i + ': could not parse expiry: ' + cellExpiry);
        // Don't reject — just skip the expiry check if unparseable
      } else if (new Date() > expires) {
        Logger.log('Token row ' + i + ': expired (was ' + cellExpiry + ')');
        return null;
      }
      
      // Get approverId — strip prefix, handle Date corruption, and try to recover
      const rawApproverIdVal = data[i][aprCol];
      let rawId;
      
      if (rawApproverIdVal instanceof Date) {
        // The cell got date-corrupted. Try to recover by matching against approvers sheet by email.
        Logger.log('Token row ' + i + ': approverId is Date, attempting recovery via email');
        const cellEmail = stripPrefix(data[i][emailCol]).toLowerCase();
        const approvers = getApprovers();
        const matched = approvers.find(a => String(a.email||'').toLowerCase() === cellEmail);
        if (matched) {
          rawId = matched.id;
          Logger.log('Recovered approverId: ' + rawId + ' via email ' + cellEmail);
        } else {
          Logger.log('Could not recover approverId for token row ' + i);
          continue;
        }
      } else {
        rawId = stripPrefix(rawApproverIdVal).toLowerCase().trim();
        // Check if the id looks like a date string — try recovery via email
        if (/^\d+\/\d+\/\d{4}/.test(rawId) || /^\d{4}-\d{2}-\d{2}/.test(rawId)) {
          Logger.log('Token row ' + i + ': approverId looks like date string (' + rawId + '), attempting recovery via email');
          const cellEmail = stripPrefix(data[i][emailCol]).toLowerCase();
          const approvers = getApprovers();
          const matched = approvers.find(a => String(a.email||'').toLowerCase() === cellEmail);
          if (matched) {
            rawId = matched.id;
            Logger.log('Recovered approverId: ' + rawId + ' via email ' + cellEmail);
          } else {
            Logger.log('Could not recover approverId via email for token row ' + i);
            continue;
          }
        }
      }
      
      Logger.log('Token VALID, approverId: ' + rawId);
      return {
        approverId:    rawId,
        approverEmail: stripPrefix(data[i][emailCol]),
        rowIndex:      i + 1,
      };
    }
  }
  Logger.log('Token NOT FOUND in any row: ' + tokenLookup);
  return null;
}

function markTokenUsed(token) {
  const sheet = getSheet(SH.TOKENS, false);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const tokenCol = data[0].indexOf('token');
  const usedCol  = data[0].indexOf('used');
  
  function stripPrefix(v){
    if (v === null || v === undefined) return '';
    let s = String(v);
    if (s.charAt(0) === "'") s = s.substring(1);
    return s;
  }
  
  const tokenLookup = String(token || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (stripPrefix(data[i][tokenCol]).trim() === tokenLookup) {
      sheet.getRange(i + 1, usedCol + 1).setValue("'true");
      return;
    }
  }
}

// ── APPROVERS ──────────────────────────────────────────────

const APR_HEADERS = ['id','name','email','role','required','defaultStatus'];

function getApprovers() {
  const rows = sheetToObjects(SH.APPROVERS);
  // Detect date-corrupted IDs (Sheets converts "apr1" to a date)
  // For each row, also keep its row index so we can reconstruct the original id
  return rows.map((r, idx) => {
    const required = r.required === true || r.required === 'true' || r.required === 'TRUE';
    let defaultStatus = String(r.defaultStatus || '').toLowerCase().trim();
    if (!['required','optional','ruleout'].includes(defaultStatus)) {
      defaultStatus = required ? 'required' : 'optional';
    }
    
    // Defensive: If id was auto-converted to a Date or date-string, reconstruct as apr<n>
    let cleanId = '';
    if (r.id instanceof Date) {
      // Use position-based reconstruction
      cleanId = 'aprv' + (idx + 1);
      Logger.log('Approver row ' + idx + ' had Date-corrupted id, using ' + cleanId);
    } else {
      cleanId = String(r.id || '').toLowerCase().trim();
      // Detect date strings like "2026-04-05" or "4/5/2026"
      if (/^\d{4}-\d{2}-\d{2}/.test(cleanId) || /^\d+\/\d+\/\d{4}/.test(cleanId)) {
        cleanId = 'aprv' + (idx + 1);
        Logger.log('Approver row ' + idx + ' had date-string id, using ' + cleanId);
      }
    }
    
    return {
      ...r,
      id:       cleanId,
      required: defaultStatus === 'required' || defaultStatus === 'ruleout',
      defaultStatus
    };
  });
}

function saveApprover(approver) {
  const sheet = getSheet(SH.APPROVERS, true);
  const data = sheet.getDataRange().getValues();
  if (!data.length || data[0][0] !== 'id') {
    sheet.clearContents();
    sheet.appendRow(APR_HEADERS);
    styleHeader(sheet, APR_HEADERS.length);
  }

  if (!approver.id) approver.id = generateId('apr');
  const current = sheet.getDataRange().getValues();
  const headers = current[0];
  const idCol = headers.indexOf('id');
  const existingIdx = current.slice(1).findIndex(r => r[idCol] === approver.id);
  const rowValues = APR_HEADERS.map(h => approver[h] !== undefined ? approver[h] : '');

  if (existingIdx >= 0) {
    sheet.getRange(existingIdx + 2, 1, 1, APR_HEADERS.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return approver;
}

function saveApprovers(approvers) {
  // Bulk replace: rewrites the entire Approvers sheet with the given array
  if (!Array.isArray(approvers)) approvers = [];
  const sheet = getSheet(SH.APPROVERS, true);
  sheet.clearContents();
  sheet.appendRow(APR_HEADERS);
  styleHeader(sheet, APR_HEADERS.length);
  approvers.forEach(a => {
    // Ensure defaultStatus + required are consistent
    let ds = String(a.defaultStatus || '').toLowerCase().trim();
    if (!['required','optional','ruleout'].includes(ds)) {
      ds = a.required ? 'required' : 'optional';
    }
    const required = (ds === 'required' || ds === 'ruleout');
    const row = [
      a.id || ('apr_' + Date.now() + Math.floor(Math.random()*1000)),
      a.name || '',
      a.email || '',
      a.role || '',
      required,
      ds
    ];
    sheet.appendRow(row);
  });
  return approvers;
}

function removeApprover(id) {
  const sheet = getSheet(SH.APPROVERS, false);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === id) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

// ── CONFIG ─────────────────────────────────────────────────

function getConfig() {
  const rows = sheetToObjects(SH.CONFIG);
  const cfg = {};
  rows.forEach(r => { cfg[r.key] = r.value; });
  return cfg;
}

/* Same as getConfig but strips sensitive keys (passwords) — used for endpoints
   that ship the config to the client. Auto-injects sheetUrl so the frontend
   can link to the sheet from the Settings page (admin-only). */
function getConfigPublic() {
  const cfg = getConfig();
  const sensitive = ['settingsPassword','appPassword','appAccessPassword'];
  const safe = {};
  Object.keys(cfg).forEach(k => {
    if (sensitive.indexOf(k) === -1) safe[k] = cfg[k];
  });
  // Inject the active spreadsheet URL so the admin link in Settings always
  // points to the CURRENT sheet (no hardcoding in HTML, no manual config).
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) safe.sheetUrl = ss.getUrl();
  } catch(e) { /* defensive */ }
  return safe;
}

function getFullConfig() {
  return {
    config:        getConfigPublic(),
    approvers:     getApprovers(),
    roleConfig:    getRoleConfig(),
    reasonOptions: getReasonOptions(),
    distLists:     getDistLists(),
    partNumbers:   getPartNumbers(),
    workCenters:   getWorkCenters(),
  };
}

/**
 * Returns whether the Settings tab gate is enabled.
 * Lock is ENABLED when Config!settingsPassword has a non-empty value.
 */
function getSettingsLockStatus() {
  const cfg = getConfig();
  const expected = String(cfg.settingsPassword || '').trim();
  return { enabled: expected.length > 0 };
}

/**
 * Verifies the Settings tab password.
 * Reads the expected password from Config!settingsPassword.
 * If empty → access granted automatically (lock disabled).
 */
function verifySettingsPassword(payload) {
  const cfg = getConfig();
  const expected = String(cfg.settingsPassword || '').trim();
  if (!expected) return { ok: true };
  const provided = (payload && typeof payload === 'object' && payload.password) ? String(payload.password) : '';
  if (!provided) return { ok: false };
  if (provided.length > 200) return { ok: false };
  const match = (provided === expected);
  if (!match) {
    Utilities.sleep(400);
  }
  return { ok: match };
}
function getAppLockStatus() {
  const cfg = getConfig();
  const expected = String(cfg.appAccessPassword || '').trim();
  return { enabled: expected.length > 0 };
}

/**
 * Verifies the app access password (the splash-screen gate on the HTML).
 * Reads the expected password from Config!appAccessPassword.
 *
 * Special behavior: if Config!appAccessPassword is EMPTY, access is granted
 * automatically (lock disabled) — the frontend should call getAppLockStatus
 * first and skip the splash entirely in that case, but we also accept any
 * input here as a defensive fallback.
 *
 * Returns: { ok: true } if match (or lock disabled), { ok: false } otherwise.
 */
function verifyAppAccess(payload) {
  const cfg = getConfig();
  const expected = String(cfg.appAccessPassword || '').trim();
  // Lock disabled when password is empty — always grant access
  if (!expected) return { ok: true };
  const provided = (payload && typeof payload === 'object' && payload.password) ? String(payload.password) : '';
  if (!provided) return { ok: false };
  // Cap input length to avoid silly inputs
  if (provided.length > 200) return { ok: false };
  const match = (provided === expected);
  if (!match) {
    // Small sleep on failure to slow down brute force attempts.
    Utilities.sleep(400);
  }
  return { ok: match };
}

function saveConfig(updates) {
  const sheet = getSheet(SH.CONFIG, true);
  const cfgHeaders = ['key', 'value'];
  const data = sheet.getDataRange().getValues();
  if (!data.length || data[0][0] !== 'key') {
    sheet.clearContents();
    sheet.appendRow(cfgHeaders);
    styleHeader(sheet, cfgHeaders.length);
  }

  const currentData = sheet.getDataRange().getValues();
  const keyCol = currentData[0].indexOf('key');

  Object.entries(updates).forEach(([key, value]) => {
    const existingIdx = currentData.slice(1).findIndex(r => r[keyCol] === key);
    if (existingIdx >= 0) {
      sheet.getRange(existingIdx + 2, 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  });
}

function getRoleConfig() {
  const rows = sheetToObjects(SH.ROLE_CONFIG);
  const rc = {};
  rows.forEach(r => { rc[r.roleKey] = r.required === true || r.required === 'true' || r.required === 'TRUE'; });
  return rc;
}

function saveRoleConfig(roleConfig) {
  const sheet = getSheet(SH.ROLE_CONFIG, true);
  const rcHeaders = ['roleKey','required'];
  sheet.clearContents();
  sheet.appendRow(rcHeaders);
  styleHeader(sheet, rcHeaders.length);
  Object.entries(roleConfig).forEach(([key, val]) => {
    sheet.appendRow([key, val]);
  });
}

function getReasonOptions() {
  const rows = sheetToObjects(SH.REASON_OPTIONS);
  return rows.map(r => ({
    ...r,
    tags: typeof r.tags === 'string' && r.tags ? JSON.parse(r.tags) : [],
  }));
}

function saveReasonOptions(options) {
  const sheet = getSheet(SH.REASON_OPTIONS, true);
  const roHeaders = ['id','label','tags'];
  sheet.clearContents();
  sheet.appendRow(roHeaders);
  styleHeader(sheet, roHeaders.length);
  options.forEach(o => {
    sheet.appendRow([o.id, o.label, JSON.stringify(o.tags || [])]);
  });
}

function getDistLists() {
  const rows = sheetToObjects(SH.DIST_LISTS);
  const lists = { creation: [], approval: [] };
  rows.forEach(r => {
    if (r.listType === 'creation') lists.creation.push(r.email);
    if (r.listType === 'approval') lists.approval.push(r.email);
  });
  return lists;
}

function saveDistLists(distLists) {
  const sheet = getSheet(SH.DIST_LISTS, true);
  const dlHeaders = ['listType','email'];
  sheet.clearContents();
  sheet.appendRow(dlHeaders);
  styleHeader(sheet, dlHeaders.length);
  ['creation','approval'].forEach(type => {
    (distLists[type] || []).forEach(email => {
      sheet.appendRow([type, email]);
    });
  });
}

// ── PART NUMBERS CATALOG ─────────────────────────────────────
function getPartNumbers() {
  let sheet = getSheet(SH.PART_NUMBERS, false);
  // Auto-create with header if missing (so users don't have to run setupSheets)
  if (!sheet) {
    sheet = getSheet(SH.PART_NUMBERS, true);
    sheet.appendRow(['partNumber','description']);
    styleHeader(sheet, 2);
    sheet.getRange(1, 1, sheet.getMaxRows(), 2).setNumberFormat('@');
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Tolerant header detection: accept several common variants of the header name
  // (case-insensitive, whitespace-insensitive). If we can't match anything,
  // fall back to using col 0 for partNumber and col 1 for description so the
  // catalog is not silently dropped if someone renames the columns by hand.
  const _norm = (h) => String(h || '').toLowerCase().replace(/[\s_\-#]/g, '');
  const headers = data[0].map(_norm);
  const PART_KEYS = ['partnumber','partno','partnum','part','sku','itemnumber','itemno'];
  const DESC_KEYS = ['description','desc','partdescription','partname','name'];
  let partCol = -1;
  for (const k of PART_KEYS) {
    const idx = headers.indexOf(k);
    if (idx >= 0) { partCol = idx; break; }
  }
  let descCol = -1;
  for (const k of DESC_KEYS) {
    const idx = headers.indexOf(k);
    if (idx >= 0) { descCol = idx; break; }
  }
  if (partCol < 0) {
    // Fallback: assume col 0 is partNumber, col 1 is description.
    Logger.log('getPartNumbers: no recognized header found in PartNumbers sheet — falling back to col 0/1. Headers were: ' + JSON.stringify(data[0]));
    partCol = 0;
    if (descCol < 0 && data[0].length > 1) descCol = 1;
  }

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = data[i][partCol];
    if (v === null || v === undefined) continue;
    const s = String(v).trim();
    if (!s) continue;
    const desc = descCol >= 0 ? String(data[i][descCol] || '').trim() : '';
    out.push({ partNumber: s, description: desc });
  }
  // Deduplicate by partNumber (case-insensitive), keep first occurrence
  const seen = {};
  return out.filter(p => {
    const k = p.partNumber.toLowerCase();
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function savePartNumbers(parts) {
  if (!Array.isArray(parts)) parts = [];
  const sheet = getSheet(SH.PART_NUMBERS, true);
  const headers = ['partNumber','description'];
  sheet.clearContents();
  sheet.appendRow(headers);
  styleHeader(sheet, headers.length);
  // Force entire columns to text format to prevent auto-conversion
  sheet.getRange(1, 1, sheet.getMaxRows(), 2).setNumberFormat('@');
  // Normalize input: accept legacy strings ('44521') OR objects ({partNumber, description})
  const seen = {};
  const clean = [];
  parts.forEach(p => {
    let pn, desc = '';
    if (typeof p === 'string') {
      pn = p.trim();
    } else if (p && typeof p === 'object') {
      pn = String(p.partNumber || '').trim();
      desc = String(p.description || '').trim();
    } else {
      return;
    }
    if (!pn) return;
    const k = pn.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    clean.push({ partNumber: pn, description: desc });
  });
  if (clean.length) {
    // Prefix both with apostrophe to force text
    const rows = clean.map(p => ["'" + p.partNumber, "'" + p.description]);
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  return clean;
}

/**
 * Upsert a single part number entry (used when a deviation form mentions
 * a part number not yet in the catalog).
 * Payload: { partNumber: '...', description: '...' }
 * - If the partNumber already exists: optionally updates description (only if non-empty AND existing was empty)
 * - If new: appends.
 */
function upsertPartNumber(payload) {
  if (!payload || typeof payload !== 'object') return null;
  const pn = String(payload.partNumber || '').trim();
  const desc = String(payload.description || '').trim();
  if (!pn) return null;
  // Length cap
  if (pn.length > 200 || desc.length > 1000) return null;

  const existing = getPartNumbers();  // returns array of {partNumber, description}
  const idxExisting = existing.findIndex(p => p.partNumber.toLowerCase() === pn.toLowerCase());
  if (idxExisting >= 0) {
    // Existing — only fill description if currently empty AND we have one to add
    if (!existing[idxExisting].description && desc) {
      existing[idxExisting].description = desc;
      savePartNumbers(existing);
    }
    return { added: false, partNumber: existing[idxExisting].partNumber, description: existing[idxExisting].description };
  } else {
    // Add new
    existing.push({ partNumber: pn, description: desc });
    savePartNumbers(existing);
    return { added: true, partNumber: pn, description: desc };
  }
}

// ── WORK CENTERS CATALOG ─────────────────────────────────────
function getWorkCenters() {
  let sheet = getSheet(SH.WORK_CENTERS, false);
  if (!sheet) {
    sheet = getSheet(SH.WORK_CENTERS, true);
    sheet.appendRow(['workCenter']);
    styleHeader(sheet, 1);
    sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Tolerant header detection: accept several common variants of the header name
  // (case-insensitive, whitespace-insensitive). If we can't match anything,
  // fall back to col 0 so the catalog is not silently dropped on header rename.
  const _norm = (h) => String(h || '').toLowerCase().replace(/[\s_\-#]/g, '');
  const headers = data[0].map(_norm);
  const WC_KEYS = ['workcenter','wc','area','workarea','center','line'];
  let wcCol = -1;
  for (const k of WC_KEYS) {
    const idx = headers.indexOf(k);
    if (idx >= 0) { wcCol = idx; break; }
  }
  if (wcCol < 0) {
    Logger.log('getWorkCenters: no recognized header found in WorkCenters sheet — falling back to col 0. Headers were: ' + JSON.stringify(data[0]));
    wcCol = 0;
  }

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = data[i][wcCol];
    if (v === null || v === undefined) continue;
    const s = String(v).trim();
    if (!s) continue;
    out.push(s);
  }
  const seen = {};
  return out.filter(w => {
    const k = w.toLowerCase();
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function saveWorkCenters(wcs) {
  if (!Array.isArray(wcs)) wcs = [];
  const sheet = getSheet(SH.WORK_CENTERS, true);
  const headers = ['workCenter'];
  sheet.clearContents();
  sheet.appendRow(headers);
  styleHeader(sheet, headers.length);
  sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  const seen = {};
  const clean = [];
  wcs.forEach(w => {
    const s = String(w || '').trim();
    if (!s) return;
    const k = s.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    clean.push(s);
  });
  if (clean.length) {
    const rows = clean.map(w => ["'" + w]);
    sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
  return clean;
}

// ── EMAIL NOTIFICATIONS ────────────────────────────────────

/**
 * Sends email notifications.
 * For 'creation': sends to ALL configured approvers with personalized token links,
 *                 plus CC to the distribution list.
 * For 'approval': sends fully-approved notice to distribution list.
 */
function sendNotification(payload) {
  const { dev, type } = payload;
  const config = getConfig();
  const distLists = getDistLists();
  const approvers = getApprovers();
  const appUrl = getWebAppUrl();

  // Always re-fetch approvals from the Sheet (don't trust the payload). This
  // matters for the edit-flow: after saveDeviation clears Approvals, the
  // frontend may still be holding the old approvals in memory. Reading from
  // the Sheet guarantees we skip only approvers who have a real, persisted
  // decision for the CURRENT cycle.
  if (dev && dev.id) {
    try {
      const approvalsRows = sheetToObjects(SH.APPROVALS);
      dev.approvals = approvalsRows.filter(a => a && a.deviationId === dev.id);
      Logger.log('sendNotification: re-fetched ' + dev.approvals.length + ' approvals for ' + dev.id + ' from Sheet (overriding payload).');
    } catch(e) {
      Logger.log('sendNotification: failed to re-fetch approvals: ' + e.message + ' (using payload as-is)');
    }
  }

  // Defensive: normalize all string-ish fields to actual strings before any .replace() happens
  // Sheets sometimes returns numbers, dates, or other types where strings are expected
  ['description','actionPlan','comments','otherReason','devNum','workCenter','initiator','riskFactor','reasonLabel','mainPartNum','date','startDate','endDate'].forEach(f => {
    if (dev[f] !== null && dev[f] !== undefined && typeof dev[f] !== 'string') {
      dev[f] = toStr(dev[f]);
    }
  });
  // Also parse parts if it came as a string
  let parts;
  try {
    let partsArr = dev.parts;
    if (typeof partsArr === 'string') {
      try { partsArr = JSON.parse(partsArr); } catch(e){ partsArr = []; }
    }
    if (!Array.isArray(partsArr)) partsArr = [];
    // Escape each partNum before joining — emails are HTML
    parts = partsArr.filter(p => p && p.partNum).map(p => escapeHtml_(p.partNum)).join(', ') || '(all)';
  } catch(e) {
    parts = '(all)';
  }

  if (type === 'creation') {
    // Log quota at start
    try {
      const remaining = MailApp.getRemainingDailyQuota();
      Logger.log('=== sendNotification(creation) START === Email quota remaining today: ' + remaining);
    } catch(e) { Logger.log('Could not check email quota: ' + e.message); }
    
    // Filter approvers by dev.selectedApprovers if present
    let targetApprovers = approvers;
    let selApprovers = dev.selectedApprovers;
    if (typeof selApprovers === 'string') {
      try { selApprovers = JSON.parse(selApprovers); } catch(e){ selApprovers = null; }
    }
    Logger.log('selectedApprovers from dev: ' + JSON.stringify(selApprovers));
    Logger.log('All approvers from sheet: ' + approvers.map(a => a.id + '/' + a.email).join(', '));
    
    if (Array.isArray(selApprovers) && selApprovers.length) {
      const selectedIds = selApprovers
        .filter(s => s && s.checked !== false)
        .map(s => String(s.id || '').toLowerCase());
      Logger.log('Selected (checked) IDs: ' + JSON.stringify(selectedIds));
      targetApprovers = approvers.filter(a => selectedIds.includes(String(a.id || '').toLowerCase()));
      Logger.log('Filtered to ' + targetApprovers.length + ' selected approvers (out of ' + approvers.length + ')');
    } else {
      Logger.log('No selectedApprovers — using ALL ' + approvers.length + ' approvers');
    }
    
    // ── Send personalized email to each approver ──
    Logger.log('=== Starting email send loop. targetApprovers count: ' + targetApprovers.length + ' ===');
    let sentCount = 0, skippedCount = 0, failedCount = 0;
    targetApprovers.forEach((approver, idx) => {
      Logger.log('[' + (idx+1) + '/' + targetApprovers.length + '] Processing approver: ' + approver.id + ' (' + (approver.email || 'NO EMAIL') + ')');
      
      // Skip approvers without email
      if (!approver.email) {
        Logger.log('  → SKIP: no email address');
        skippedCount++;
        return;
      }
      // Check if already approved
      const alreadyDecided = (dev.approvals || []).find(a => String(a.approverId || '').toLowerCase() === String(approver.id || '').toLowerCase());
      if (alreadyDecided) {
        Logger.log('  → SKIP: approver already submitted decision (decision=' + alreadyDecided.decision + ')');
        skippedCount++;
        return;
      }

      try {
        Logger.log('  → Creating token for ' + approver.id + '...');
        const token = createApproverToken(approver.id, approver.email);
        Logger.log('  → Token created: ' + token.substring(0, 8) + '...');
        const approvalLink = appUrl + '?token=' + token;
        const subject = `[Action Required] Deviation ${dev.devNum} – WC ${dev.workCenter} needs your signature`;
        const htmlBody = buildCreationEmailHtml(dev, approver, approvalLink, parts, config);

        _sendEmailSafe(approver.email, subject, htmlBody);
        Logger.log('  → Email SENT successfully to ' + approver.email);
        sentCount++;
      } catch(err) {
        Logger.log('  → FAILED to send to ' + approver.email + ': ' + err.message + ' | stack: ' + (err.stack || 'no stack'));
        failedCount++;
      }
      
      // Small delay to avoid quota throttling
      Utilities.sleep(200);
    });
    Logger.log('=== Email send loop complete: sent=' + sentCount + ', skipped=' + skippedCount + ', failed=' + failedCount + ' ===');
    try {
      Logger.log('Email quota remaining after loop: ' + MailApp.getRemainingDailyQuota());
    } catch(e) {}

    // CC distribution list (informational, no action link)
    const distEmails = distLists.creation || [];
    Logger.log('=== FYI distribution list: ' + JSON.stringify(distEmails) + ' ===');
    if (distEmails.length) {
      const subject = `[FYI] New Deviation Submitted: ${dev.devNum} – WC ${dev.workCenter}`;
      const htmlBody = buildCreationDistEmailHtml(dev, parts, config);
      Logger.log('Sending FYI to: TO=' + distEmails[0] + ', CC=' + distEmails.slice(1).join(','));
      try {
        MailApp.sendEmail({
          to:       distEmails[0],
          cc:       distEmails.slice(1).join(','),
          subject:  subject,
          htmlBody: htmlBody,
          name:     'MWAAF Deviation System',
        });
        Logger.log('FYI email sent successfully');
      } catch(err) {
        Logger.log('Failed to send dist email: ' + err.message);
      }
    } else {
      Logger.log('No FYI distribution list configured (DistLists sheet rows with listType=creation)');
    }

  } else if (type === 'approval') {
    // ── Send fully-approved notice ──
    // Only to the "approval" Distribution List (not to approvers).
    const approvalEmails = [...new Set(distLists.approval || [])].filter(e => e);
    Logger.log('=== Approval notification: dist list "approval" = ' + JSON.stringify(approvalEmails) + ' ===');

    if (approvalEmails.length) {
      const subject = `✅ Deviation APPROVED: ${dev.devNum} – WC ${dev.workCenter}`;
      const htmlBody = buildApprovalEmailHtml(dev, parts, config);
      try {
        MailApp.sendEmail({
          to:       approvalEmails[0],
          cc:       approvalEmails.slice(1).join(','),
          subject:  subject,
          htmlBody: htmlBody,
          name:     'MWAAF Deviation System',
        });
        Logger.log('Approval email sent to dist list (' + approvalEmails.length + ' recipients)');
      } catch(err) {
        Logger.log('Failed to send approval email: ' + err.message);
      }
    } else {
      Logger.log('No "approval" distribution list configured. Skipping approval email.');
    }

  } else if (type === 'rejected') {
    // ── Send rejected notice ──
    // Only to the "approval" Distribution List (same recipients as on-approval).
    const approvalEmails = [...new Set(distLists.approval || [])].filter(e => e);
    Logger.log('=== Rejected notification: dist list "approval" = ' + JSON.stringify(approvalEmails) + ' ===');

    if (approvalEmails.length) {
      const subject = `❌ Deviation REJECTED: ${dev.devNum} – WC ${dev.workCenter}`;
      const htmlBody = buildRejectedEmailHtml(dev, parts, config);
      try {
        MailApp.sendEmail({
          to:       approvalEmails[0],
          cc:       approvalEmails.slice(1).join(','),
          subject:  subject,
          htmlBody: htmlBody,
          name:     'MWAAF Deviation System',
        });
        Logger.log('Rejected email sent to dist list (' + approvalEmails.length + ' recipients)');
      } catch(err) {
        Logger.log('Failed to send rejected email: ' + err.message);
      }
    } else {
      Logger.log('No "approval" distribution list configured. Skipping rejected email.');
    }
  }

  return { sent: true };
}

// ── EMAIL HTML BUILDERS ────────────────────────────────────

function emailWrapper(content) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body { font-family: Calibri, Arial, sans-serif; background: #f4f7f0; margin: 0; padding: 20px; }
    .card { background: #fff; border-radius: 10px; max-width: 600px; margin: 0 auto;
            border: 1.5px solid #d4e0b8; box-shadow: 0 2px 12px rgba(90,120,40,.10); overflow: hidden; }
    .hdr { background: linear-gradient(135deg, #7A9A3A, #39B54A); padding: 24px 28px; }
    .hdr-title { color: #fff; font-size: 22px; font-weight: 700; margin: 0; }
    .hdr-sub { color: rgba(255,255,255,.85); font-size: 13px; margin-top: 4px; }
    .body { padding: 24px 28px; }
    .field { margin-bottom: 14px; }
    .label { font-size: 10px; font-weight: 700; color: #7A9A3A; text-transform: uppercase;
             letter-spacing: .6px; margin-bottom: 3px; }
    .val { font-size: 15px; color: #2c3e1a; font-weight: 600; }
    .desc { font-size: 13px; color: #4a5e30; line-height: 1.6; background: #f2f7e8;
            border-left: 3px solid #7A9A3A; padding: 10px 14px; border-radius: 0 6px 6px 0; }
    .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 14px; }
    .btn { display: inline-block; background: #007b8a; color: #fff; text-decoration: none;
           padding: 14px 32px; border-radius: 8px; font-size: 16px; font-weight: 700;
           text-align: center; margin: 10px 0; }
    .btn-sm { background: #7A9A3A; padding: 10px 22px; font-size: 14px; }
    .footer { background: #f8faf4; border-top: 1.5px solid #d4e0b8; padding: 14px 28px;
              font-size: 11px; color: #7a8f5c; }
    .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 11px;
             font-weight: 700; }
    .badge-pending { background: #fff8e1; color: #c97c00; border: 1px solid #f0c050; }
    .badge-required { background: #e6f6f8; color: #005e6b; border: 1px solid #007b8a; }
    .badge-optional { background: #f2f7e8; color: #5e7a27; border: 1px solid #7A9A3A; }
  </style></head><body>
  <div class="card">${content}</div>
  </body></html>`;
}

// ── HTML SAFETY (XSS prevention for emails and approver view) ──
function escapeHtml_(v) {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
// Multi-line escape: escape HTML then convert newlines to <br>
function escapeHtmlNl_(v) {
  return escapeHtml_(v).replace(/\n/g, '<br>');
}

function buildCreationEmailHtml(dev, approver, approvalLink, parts, config) {
  const requiredBadge = approver.required
    ? '<span class="badge badge-required">🔒 Required Approver</span>'
    : '<span class="badge badge-optional">🔓 Optional Approver</span>';

  const content = `
    <div class="hdr">
      <div class="hdr-title">🔔 Deviation Approval Required</div>
      <div class="hdr-sub">MWAAF Deviation Authorization System · QF158</div>
    </div>
    <div class="body">
      <p style="margin:0 0 16px;font-size:14px;color:#2c3e1a">
        Hello <strong>${escapeHtml_(approver.name)}</strong>, a new deviation has been submitted and requires your review and signature.
      </p>
      <p style="margin:0 0 18px;font-size:13px;color:#4a5e30">
        Your role: <strong>${escapeHtml_(approver.role)}</strong> &nbsp; ${requiredBadge}
      </p>

      <div class="grid2">
        <div class="field"><div class="label">Deviation #</div><div class="val" style="color:#7A9A3A">${escapeHtml_(dev.devNum)}</div></div>
        <div class="field"><div class="label">Status</div><div class="val"><span class="badge badge-pending">⏳ Pending Approval</span></div></div>
        <div class="field"><div class="label">Initiator</div><div class="val">${escapeHtml_(dev.initiator) || '—'}</div></div>
        <div class="field"><div class="label">Work Center</div><div class="val">${escapeHtml_(dev.workCenter) || '—'}</div></div>
        <div class="field"><div class="label">Start Date</div><div class="val">${escapeHtml_(fmtUS(dev.startDate)) || '—'}</div></div>
        <div class="field"><div class="label">End Date</div><div class="val">${escapeHtml_(fmtUS(dev.endDate)) || '—'}</div></div>
        <div class="field"><div class="label">Part(s)</div><div class="val">${parts}</div></div>
      </div>

      <div class="field">
        <div class="label">Description</div>
        <div class="desc">${escapeHtmlNl_(dev.description)}</div>
      </div>
      <div class="field">
        <div class="label">Action Plan</div>
        <div class="desc">${escapeHtmlNl_(dev.actionPlan)}</div>
      </div>

      <div style="text-align:center;padding:20px 0 10px">
        <p style="font-size:13px;color:#4a5e30;margin-bottom:14px">
          Click the button below to open <strong>your personalized approval view</strong>.<br>
          You will only see deviations pending <em>your</em> signature.
        </p>
        <a href="${escapeHtml_(approvalLink)}" class="btn">✅ Review &amp; Sign Deviation</a>
        <p style="font-size:11px;color:#7a8f5c;margin-top:10px">
          This link is personal and valid for 7 days. Do not share it.
        </p>
      </div>
    </div>
    <div class="footer">
      MWAAF Deviation Authorization System · QF158 · Rev.1<br>
      This email was generated automatically. Do not reply directly to this email.
    </div>`;
  return emailWrapper(content);
}

function buildCreationDistEmailHtml(dev, parts, config) {
  // Build a link to the read-only deviation view (same as approval-complete email)
  const appUrl = (config && config.appUrl) || ScriptApp.getService().getUrl() || '';
  const viewLink = appUrl ? appUrl + (appUrl.indexOf('?') >= 0 ? '&' : '?') + 'view=' + encodeURIComponent(dev.devNum || dev.id || '') : '';
  const linkHtml = viewLink ? `
      <div style="text-align:center;padding:20px 0 10px">
        <a href="${escapeHtml_(viewLink)}" style="display:inline-block;background:#7A9A3A;color:#fff;text-decoration:none;padding:11px 26px;border-radius:8px;font-weight:700;font-size:14px;border:1px solid #5e7a27">📋 View Deviation Details</a>
        <p style="font-size:11px;color:#7a8f5c;margin-top:10px">Click the button above to view this deviation in read-only mode (no login required).</p>
      </div>` : '';

  const content = `
    <div class="hdr">
      <div class="hdr-title">📋 New Deviation Submitted</div>
      <div class="hdr-sub">MWAAF Deviation Authorization System · QF158 — For your information</div>
    </div>
    <div class="body">
      <div class="grid2">
        <div class="field"><div class="label">Deviation #</div><div class="val" style="color:#7A9A3A">${escapeHtml_(dev.devNum)}</div></div>
        <div class="field"><div class="label">Initiator</div><div class="val">${escapeHtml_(dev.initiator) || '—'}</div></div>
        <div class="field"><div class="label">Work Center</div><div class="val">${escapeHtml_(dev.workCenter) || '—'}</div></div>
        <div class="field"><div class="label">Part(s)</div><div class="val">${parts}</div></div>
        <div class="field"><div class="label">Start Date</div><div class="val">${escapeHtml_(fmtUS(dev.startDate)) || '—'}</div></div>
        <div class="field"><div class="label">End Date</div><div class="val">${escapeHtml_(fmtUS(dev.endDate)) || '—'}</div></div>
      </div>
      <div class="field"><div class="label">Description</div>
        <div class="desc">${escapeHtmlNl_(dev.description)}</div></div>
      ${linkHtml}
    </div>
    <div class="footer">This is an informational copy. Approvers have been notified separately.</div>`;
  return emailWrapper(content);
}

function buildApprovalEmailHtml(dev, parts, config) {
  // Build a link to the read-only deviation view
  const appUrl = (config && config.appUrl) || ScriptApp.getService().getUrl() || '';
  const viewLink = appUrl ? appUrl + (appUrl.indexOf('?') >= 0 ? '&' : '?') + 'view=' + encodeURIComponent(dev.devNum || dev.id || '') : '';
  const linkHtml = viewLink ? `
      <div style="text-align:center;padding:20px 0 10px">
        <a href="${escapeHtml_(viewLink)}" style="display:inline-block;background:#2a9438;color:#fff;text-decoration:none;padding:11px 26px;border-radius:8px;font-weight:700;font-size:14px;border:1px solid #1f7a2c">📋 View Deviation Details</a>
        <p style="font-size:11px;color:#7a8f5c;margin-top:10px">Click the button above to view this deviation in read-only mode (no login required).</p>
      </div>` : '';

  const content = `
    <div class="hdr" style="background:linear-gradient(135deg,#2a9438,#39B54A)">
      <div class="hdr-title">✅ Deviation Fully Approved &amp; Active</div>
      <div class="hdr-sub">MWAAF Deviation Authorization System · QF158</div>
    </div>
    <div class="body">
      <p style="margin:0 0 16px;font-size:14px;color:#2c3e1a">
        Deviation <strong>${escapeHtml_(dev.devNum)}</strong> has received all required signatures and is now <strong>ACTIVE</strong>.
      </p>
      <div class="grid2">
        <div class="field"><div class="label">Deviation #</div><div class="val" style="color:#2a9438">${escapeHtml_(dev.devNum)}</div></div>
        <div class="field"><div class="label">Initiator</div><div class="val">${escapeHtml_(dev.initiator) || '—'}</div></div>
        <div class="field"><div class="label">Work Center</div><div class="val">${escapeHtml_(dev.workCenter) || '—'}</div></div>
        <div class="field"><div class="label">Part(s)</div><div class="val">${parts}</div></div>
        <div class="field"><div class="label">Start Date</div><div class="val">${escapeHtml_(fmtUS(dev.startDate)) || '—'}</div></div>
        <div class="field"><div class="label">End Date</div><div class="val">${escapeHtml_(fmtUS(dev.endDate)) || '—'}</div></div>
      </div>
      <div class="field"><div class="label">Description</div>
        <div class="desc">${escapeHtmlNl_(dev.description)}</div></div>
      ${linkHtml}
    </div>
    <div class="footer">MWAAF Deviation Authorization System · QF158 · Rev.1</div>`;
  return emailWrapper(content);
}

function buildRejectedEmailHtml(dev, parts, config) {
  // Build a link to the read-only deviation view
  const appUrl = (config && config.appUrl) || ScriptApp.getService().getUrl() || '';
  const viewLink = appUrl ? appUrl + (appUrl.indexOf('?') >= 0 ? '&' : '?') + 'view=' + encodeURIComponent(dev.devNum || dev.id || '') : '';
  const linkHtml = viewLink ? `
      <div style="text-align:center;padding:20px 0 10px">
        <a href="${escapeHtml_(viewLink)}" style="display:inline-block;background:#c0392b;color:#fff;text-decoration:none;padding:11px 26px;border-radius:8px;font-weight:700;font-size:14px;border:1px solid #962d22">📋 View Deviation Details</a>
        <p style="font-size:11px;color:#7a8f5c;margin-top:10px">Click the button above to view this deviation in read-only mode (no login required).</p>
      </div>` : '';

  // Find the rejecting approver(s) from approvals array
  let rejectionInfo = '';
  try {
    const approvals = Array.isArray(dev.approvals) ? dev.approvals : [];
    const rejections = approvals.filter(a => a && a.decision === 'rejected');
    if (rejections.length) {
      const rows = rejections.map(r => {
        const who = escapeHtml_(r.approverName || r.approverId || '(unknown)');
        const role = escapeHtml_(r.approverRole || '');
        const cmt = r.comments ? escapeHtmlNl_(r.comments) : '<i style="color:#999">(no comments)</i>';
        return `<div style="background:#fff5f4;border-left:3px solid #c0392b;padding:10px 14px;border-radius:0 6px 6px 0;margin-bottom:8px">
          <div style="font-size:13px;font-weight:700;color:#962d22">${who}${role ? ' · <span style="font-weight:400;font-size:12px">'+role+'</span>' : ''}</div>
          <div style="font-size:12px;color:#4a5e30;margin-top:4px;line-height:1.5">${cmt}</div>
        </div>`;
      }).join('');
      rejectionInfo = `
        <div class="field"><div class="label" style="color:#c0392b">Rejection Details</div>
          ${rows}
        </div>`;
    }
  } catch(e) { /* defensive */ }

  const content = `
    <div class="hdr" style="background:linear-gradient(135deg,#962d22,#c0392b)">
      <div class="hdr-title">❌ Deviation Rejected</div>
      <div class="hdr-sub">MWAAF Deviation Authorization System · QF158</div>
    </div>
    <div class="body">
      <p style="margin:0 0 16px;font-size:14px;color:#2c3e1a">
        Deviation <strong>${escapeHtml_(dev.devNum)}</strong> has been <strong>REJECTED</strong> by a required approver. The deviation is not authorized to proceed.
      </p>
      <div class="grid2">
        <div class="field"><div class="label">Deviation #</div><div class="val" style="color:#c0392b">${escapeHtml_(dev.devNum)}</div></div>
        <div class="field"><div class="label">Initiator</div><div class="val">${escapeHtml_(dev.initiator) || '—'}</div></div>
        <div class="field"><div class="label">Work Center</div><div class="val">${escapeHtml_(dev.workCenter) || '—'}</div></div>
        <div class="field"><div class="label">Part(s)</div><div class="val">${parts}</div></div>
        <div class="field"><div class="label">Start Date</div><div class="val">${escapeHtml_(fmtUS(dev.startDate)) || '—'}</div></div>
        <div class="field"><div class="label">End Date</div><div class="val">${escapeHtml_(fmtUS(dev.endDate)) || '—'}</div></div>
      </div>
      <div class="field"><div class="label">Description</div>
        <div class="desc">${escapeHtmlNl_(dev.description)}</div></div>
      ${rejectionInfo}
      ${linkHtml}
    </div>
    <div class="footer">MWAAF Deviation Authorization System · QF158 · Rev.1</div>`;
  return emailWrapper(content);
}

// ── APPROVER VIEW (served from token link) ─────────────────

/**
 * Serves the personalized ApproverView HTML page.
 * Validates token and injects approver data + pending deviations.
 */
function serveApproverView(token) {
  Logger.log('serveApproverView called with token: ' + token);
  
  const tokenData = validateToken(token);
  Logger.log('tokenData: ' + JSON.stringify(tokenData));

  if (!tokenData) {
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><style>body{font-family:Calibri,sans-serif;display:flex;' +
      'align-items:center;justify-content:center;min-height:100vh;background:#f4f7f0;margin:0}' +
      '.card{background:#fff;border-radius:12px;padding:36px;text-align:center;max-width:400px;' +
      'border:1.5px solid #d4e0b8}</style></head><body><div class="card">' +
      '<h2 style="color:#c0392b">&#10060; Link Invalid or Expired</h2>' +
      '<p style="color:#7a8f5c;font-size:14px">This link has expired, already been used, or is invalid.</p>' +
      '<p style="color:#7a8f5c;font-size:14px">Please ask for a new deviation notification email.</p>' +
      '</div></body></html>'
    ).setTitle('Invalid Link');
  }

  // Get approver info
  const approvers = getApprovers();
  Logger.log('All approvers: ' + JSON.stringify(approvers.map(a => ({id:a.id, name:a.name}))));
  Logger.log('Looking for approverId: ' + tokenData.approverId);
  
  const approver = approvers.find(a => a.id === tokenData.approverId);
  Logger.log('Found approver: ' + JSON.stringify(approver));
  
  if (!approver) {
    // Show debug info instead of blank error
    const allIds = approvers.map(a => a.id).join(', ');
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><style>body{font-family:Calibri,sans-serif;padding:30px;' +
      'max-width:600px;margin:0 auto;background:#f4f7f0}</style></head><body>' +
      '<h2 style="color:#c0392b">Approver Not Found</h2>' +
      '<p><b>Token approverId:</b> <code>' + tokenData.approverId + '</code></p>' +
      '<p><b>Available IDs in sheet:</b> <code>' + allIds + '</code></p>' +
      '<p style="color:#7a8f5c">The ID in the token does not match any approver in the sheet. ' +
      'Please update the Approvers sheet so the id column matches, then resend the notification.</p>' +
      '</body></html>'
    ).setTitle('Approver Not Found');
  }

  // Get ALL deviations involving this approver (pending + decided), for the side panel
  const allDevs = getDeviations({});
  const aprIdLc = String(approver.id || '').toLowerCase();
  const relevantDevs = allDevs.filter(function(d) {
    // Pending: this approver hasn't decided yet AND status is pending/partial
    const myDecision = (d.approvals || []).find(function(a) {
      return String(a.approverId || '').toLowerCase() === aprIdLc;
    });
    // Include if approver has decided OR if pending/partial (waiting for decision)
    return myDecision || d.status === 'pending' || d.status === 'partial' || d.status === 'rejected' || d.status === 'approved';
  });
  
  // For now keep all deviations from sheet (pending + decided) so left panel can list them
  Logger.log('Relevant devs for ' + approver.id + ': ' + relevantDevs.length);

  // Keep photos (base64 dataUrls) so approver can review evidence
  const devsForView = relevantDevs.map(function(d) {
    return JSON.parse(JSON.stringify(d));
  });

  // Build HTML entirely in Code.gs — no dependency on ApproverView.html file cache
  const html = buildApproverViewHtml(approver, devsForView, token);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Pending Approvals – ' + approver.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Serves a read-only HTML view of a single deviation by devNum (e.g. "DEV-1004").
 * Linked from approval-complete emails. No authentication required because the URL
 * contains only a publicly-known deviation number — sensitive operations are NOT exposed.
 * (The approver view, which CAN take action, requires a token.)
 */
function serveReadOnlyDeviationView(devNumOrId) {
  const requested = String(devNumOrId || '').trim();
  if (!requested) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:Calibri,sans-serif;padding:30px;text-align:center"><h2 style="color:#c0392b">&#10060; Invalid request</h2><p>No deviation specified.</p></body></html>'
    ).setTitle('Not Found');
  }
  // Look up by devNum (preferred) or id
  const allDevs = sheetToObjects(SH.DEVIATIONS);
  const dev = allDevs.find(d =>
    String(d.devNum || '').toLowerCase() === requested.toLowerCase() ||
    String(d.id || '').toLowerCase() === requested.toLowerCase()
  );
  if (!dev) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:Calibri,sans-serif;padding:30px;text-align:center;background:#f4f7f0">' +
      '<h2 style="color:#c0392b">&#10060; Deviation not found</h2>' +
      '<p style="color:#7a8f5c">"' + escapeHtml_(requested) + '" was not found in the system.</p>' +
      '</body></html>'
    ).setTitle('Not Found');
  }
  // Hydrate parts/photos/etc. (sheetToObjects gives us strings for JSON columns)
  ['parts','reasons','tags','fourm','photos','selectedApprovers'].forEach(f => {
    if (typeof dev[f] === 'string') {
      try { dev[f] = JSON.parse(dev[f]); } catch(e){ dev[f] = []; }
    }
  });
  // Attach approvals
  const approvals = sheetToObjects(SH.APPROVALS).filter(a => a.deviationId === dev.id);
  dev.approvals = approvals;

  const html = buildReadOnlyDeviationHtml(dev);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Deviation ' + (dev.devNum || dev.id))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Builds an HTML page showing one deviation in read-only format with all details:
 * meta, description, parts, reason, action plan, photos, and approval signatures.
 * No interactive controls.
 */
function buildReadOnlyDeviationHtml(dev) {
  function esc(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  }
  function escNl(v) { return esc(v).replace(/\n/g, '<br>'); }
  function fmtUS_(d) {
    if (!d) return '—';
    try {
      const date = (d instanceof Date) ? d : new Date(d);
      if (isNaN(date)) return esc(d);
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      const yy = String(date.getFullYear()).slice(-2);
      return m + '/' + dd + '/' + yy;
    } catch (e) { return esc(d); }
  }
  function fmtDateTimeUS_(d) {
    if (!d) return '—';
    try {
      const date = (d instanceof Date) ? d : new Date(d);
      if (isNaN(date)) return esc(d);
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      const yy = String(date.getFullYear()).slice(-2);
      const hh = String(date.getHours()).padStart(2, '0');
      const mn = String(date.getMinutes()).padStart(2, '0');
      return m + '/' + dd + '/' + yy + ' ' + hh + ':' + mn;
    } catch(e) { return esc(d); }
  }

  const status = String(dev.status || 'draft');
  const statusLabel = {pending:'Pending',partial:'Partial',approved:'Approved',rejected:'Rejected',expired:'Expired',draft:'Draft'}[status] || status;
  const statusColor = status === 'approved' ? '#2a9438' : status === 'rejected' ? '#c0392b' : '#c97c00';

  // Parts table
  const partsArr = Array.isArray(dev.parts) ? dev.parts : [];
  const partsValid = partsArr.filter(p => p && (p.partNum || p.partName));
  let partsHtml = '';
  if (partsValid.length) {
    partsHtml = '<table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:#5e7a27;color:#fff"><th style="padding:8px;text-align:left;border:1px solid #5e7a27">Part #</th><th style="padding:8px;text-align:left;border:1px solid #5e7a27">Qty</th><th style="padding:8px;text-align:left;border:1px solid #5e7a27">Part Name</th></tr></thead><tbody>';
    partsValid.forEach(p => {
      partsHtml += '<tr><td style="padding:6px 8px;border:1px solid #d4e0b8">' + esc(p.partNum || '—') + '</td><td style="padding:6px 8px;border:1px solid #d4e0b8">' + esc(p.qty || '—') + '</td><td style="padding:6px 8px;border:1px solid #d4e0b8">' + esc(p.partName || '—') + '</td></tr>';
    });
    partsHtml += '</tbody></table>';
  } else {
    partsHtml = '<div style="color:#7a8f5c;font-style:italic">None specified</div>';
  }

  // Photos
  const photos = (Array.isArray(dev.photos) ? dev.photos : [])
    .filter(p => p && p.dataUrl && /^data:image\//i.test(String(p.dataUrl)));
  let photosHtml = '';
  if (photos.length) {
    photosHtml = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:10px">';
    photos.forEach((p, idx) => {
      photosHtml += '<div onclick="openPhotoLightbox(' + idx + ')" title="Click to enlarge" style="border:1.5px solid #d4e0b8;border-radius:8px;overflow:hidden;background:#f4f7f0;display:flex;align-items:center;justify-content:center;min-height:140px;max-height:200px;cursor:zoom-in"><img src="' + esc(p.dataUrl) + '" style="max-width:100%;max-height:200px;object-fit:contain;display:block;pointer-events:none"></div>';
    });
    photosHtml += '</div>';
  }

  // Approvals
  const approvers = getApprovers();
  const aprByIdLookup = {};
  approvers.forEach(a => { aprByIdLookup[String(a.id || '').toLowerCase()] = a; });

  let selectedApprovers = [];
  if (dev.selectedApprovers) {
    try {
      selectedApprovers = typeof dev.selectedApprovers === 'string' ? JSON.parse(dev.selectedApprovers) : dev.selectedApprovers;
    } catch(e) { selectedApprovers = []; }
  }
  selectedApprovers = (selectedApprovers || []).filter(s => s && s.checked !== false);
  if (!selectedApprovers.length) {
    selectedApprovers = approvers.map(a => ({
      id: a.id,
      checked: true,
      status: a.defaultStatus || (a.required ? 'required' : 'optional')
    }));
  }

  let approvalsHtml = '<table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:#5e7a27;color:#fff">' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Approver</th>' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Role</th>' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Status</th>' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Decision</th>' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Date</th>' +
    '<th style="padding:8px;text-align:left;border:1px solid #5e7a27">Comments</th>' +
    '</tr></thead><tbody>';
  selectedApprovers.forEach(s => {
    const aprData = aprByIdLookup[String(s.id || '').toLowerCase()];
    if (!aprData) return;
    const ap = (dev.approvals || []).find(a => String(a.approverId || '').toLowerCase() === String(s.id || '').toLowerCase());
    const statusLbl = {required:'🔒 Required',optional:'🔓 Optional',ruleout:'⚡ Rule-Out'}[s.status] || s.status;
    const decisionLbl = ap ? (ap.decision === 'approved' ? '<span style="color:#2a9438;font-weight:700">✅ Approved</span>' : '<span style="color:#c0392b;font-weight:700">❌ Rejected</span>') : '<span style="color:#c97c00">Pending</span>';
    approvalsHtml += '<tr>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8"><strong>' + esc(aprData.name) + '</strong><br><span style="font-size:11px;color:#7a8f5c">' + esc(aprData.email) + '</span></td>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8">' + esc(aprData.role) + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8">' + statusLbl + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8">' + decisionLbl + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8">' + (ap ? fmtDateTimeUS_(ap.date) : '—') + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #d4e0b8">' + (ap && ap.comments ? esc(ap.comments) : '—') + '</td>' +
      '</tr>';
  });
  approvalsHtml += '</tbody></table>';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Deviation ' + esc(dev.devNum) + '</title>' +
    '<style>' +
    'body{font-family:Calibri,Arial,sans-serif;background:#f4f7f0;margin:0;padding:20px;color:#2c3e1a}' +
    '.wrap{max-width:900px;margin:0 auto;background:#fff;border:1.5px solid #d4e0b8;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.05)}' +
    '.hdr{background:linear-gradient(135deg,#5e7a27,#7A9A3A);color:#fff;padding:20px 24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}' +
    '.hdr-title{font-size:20px;font-weight:800;margin:0}' +
    '.hdr-sub{font-size:12px;opacity:.85;margin-top:2px}' +
    '.dev-id{font-size:24px;font-weight:800}' +
    '.status{display:inline-block;padding:4px 12px;border-radius:6px;background:#fff;font-size:13px;font-weight:700;margin-top:4px}' +
    '.body{padding:24px}' +
    '.section{margin-bottom:24px}' +
    '.section-title{font-size:11px;font-weight:700;color:#7a8f5c;text-transform:uppercase;margin-bottom:8px;letter-spacing:.5px}' +
    '.meta-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:0;border:1.5px solid #d4e0b8;border-radius:8px;overflow:hidden;margin-bottom:24px}' +
    '.meta-cell{padding:10px 14px;border-right:1px solid #d4e0b8;border-bottom:1px solid #d4e0b8}' +
    '.meta-cell:nth-child(4n){border-right:none}' +
    '.meta-cell:nth-last-child(-n+4){border-bottom:none}' +
    '.meta-label{font-size:10px;font-weight:700;color:#7a8f5c;text-transform:uppercase;letter-spacing:.4px}' +
    '.meta-val{font-size:13px;font-weight:600;color:#2c3e1a;margin-top:2px}' +
    '.text-block{background:#f8faf4;border:1px solid #d4e0b8;border-radius:6px;padding:12px 14px;font-size:14px;line-height:1.5}' +
    '.footer{background:#f8faf4;padding:14px 24px;font-size:11px;color:#7a8f5c;text-align:center;border-top:1px solid #d4e0b8}' +
    '@media (max-width:600px){.meta-grid{grid-template-columns:1fr 1fr}.meta-cell:nth-child(4n){border-right:1px solid #d4e0b8}.meta-cell:nth-child(2n){border-right:none}}' +
    '</style></head><body>' +
    '<div class="wrap">' +
    '<div class="hdr"><div><div class="hdr-title">📋 Deviation View</div><div class="hdr-sub">MWAAF Deviation Authorization System · QF158 · Read-only</div></div>' +
    '<div style="text-align:right"><div class="dev-id">' + esc(dev.devNum) + '</div><div class="status" style="color:' + statusColor + '">' + esc(statusLabel) + '</div></div></div>' +
    '<div class="body">' +
    '<div class="meta-grid">' +
      '<div class="meta-cell"><div class="meta-label">Date</div><div class="meta-val">' + fmtUS_(dev.date) + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Submitted</div><div class="meta-val">' + fmtDateTimeUS_(dev.submittedAt) + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Part Number</div><div class="meta-val">' + esc(dev.mainPartNum || '—') + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Work Center</div><div class="meta-val">' + esc(dev.workCenter || '—') + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Initiator</div><div class="meta-val">' + esc(dev.initiator || '—') + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Cust. Approval</div><div class="meta-val">' + esc(dev.custApproval || 'No') + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">Start Date</div><div class="meta-val">' + fmtUS_(dev.startDate) + '</div></div>' +
      '<div class="meta-cell"><div class="meta-label">End Date</div><div class="meta-val">' + fmtUS_(dev.endDate) + '</div></div>' +
    '</div>' +
    '<div class="section"><div class="section-title">Description</div><div class="text-block">' + (escNl(dev.description) || '—') + '</div></div>' +
    '<div class="section"><div class="section-title">Affected Parts</div>' + partsHtml + '</div>' +
    '<div class="section"><div class="section-title">Reason</div><div class="text-block"><strong>' + esc(dev.reasonLabel || (Array.isArray(dev.reasons) && dev.reasons[0]) || 'None') + '</strong>' +
      ((Array.isArray(dev.fourm) && dev.fourm.length) ? '<div style="margin-top:6px;font-size:12px;color:#7a8f5c">4M: ' + esc(dev.fourm.join(', ')) + '</div>' : '') +
      (dev.otherReason ? '<div style="margin-top:6px;font-size:12px"><strong>Other:</strong> ' + esc(dev.otherReason) + '</div>' : '') +
      '<div style="margin-top:6px;font-size:12px"><strong>Risk:</strong> ' + esc(dev.riskFactor || 'Low') + '</div>' +
    '</div></div>' +
    '<div class="section"><div class="section-title">Action Plan</div><div class="text-block">' + (escNl(dev.actionPlan) || '—') + (dev.comments ? '<div style="margin-top:8px;font-size:12px;color:#7a8f5c"><strong>Comments:</strong> ' + esc(dev.comments) + '</div>' : '') + '</div></div>' +
    (photos.length ? '<div class="section"><div class="section-title">Photo Evidence (' + photos.length + ')</div>' + photosHtml + '</div>' : '') +
    '<div class="section"><div class="section-title">Approvals</div>' + approvalsHtml + '</div>' +
    '</div>' +
    '<div class="footer">MWAAF Deviation Authorization System · QF158 · This is a read-only view.</div>' +
    '</div>' +
    /* Lightbox overlay (hidden by default) */
    '<div id="photoLightbox" onclick="closePhotoLightbox()" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.92);z-index:9999;align-items:center;justify-content:center;padding:20px;cursor:zoom-out">' +
      '<div style="position:relative;max-width:95vw;max-height:95vh">' +
        '<button onclick="event.stopPropagation();closePhotoLightbox()" style="position:absolute;top:-40px;right:0;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:14px;font-weight:600;font-family:inherit">&#x2715; Close</button>' +
        '<img id="photoLightboxImg" src="" style="max-width:95vw;max-height:90vh;display:block;border-radius:8px;box-shadow:0 12px 40px rgba(0,0,0,0.5)">' +
        '<div id="photoLightboxLabel" style="text-align:center;color:rgba(255,255,255,0.85);font-size:12px;margin-top:8px;font-family:Calibri,sans-serif"></div>' +
      '</div>' +
    '</div>' +
    /* Inline photo data + lightbox script */
    '<script>' +
      'var _PHOTOS = ' + JSON.stringify(photos.map(p => p.dataUrl)) + ';' +
      'function openPhotoLightbox(idx){' +
        'if (idx < 0 || idx >= _PHOTOS.length) return;' +
        'var lb = document.getElementById("photoLightbox");' +
        'var img = document.getElementById("photoLightboxImg");' +
        'var lbl = document.getElementById("photoLightboxLabel");' +
        'if (!lb || !img) return;' +
        'img.src = _PHOTOS[idx];' +
        'if (lbl) lbl.textContent = "Image " + (idx+1) + " of " + _PHOTOS.length;' +
        'lb.style.display = "flex";' +
        'document.body.style.overflow = "hidden";' +
      '}' +
      'function closePhotoLightbox(){' +
        'var lb = document.getElementById("photoLightbox");' +
        'if (lb) lb.style.display = "none";' +
        'document.body.style.overflow = "";' +
      '}' +
      'document.addEventListener("keydown", function(e){ if (e.key === "Escape") closePhotoLightbox(); });' +
    '</script>' +
    '</body></html>';
}

function buildApproverViewHtml(approver, deviations, token) {
  function esc(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  var aprIdLc = String(approver.id || '').toLowerCase();

  // ── Build sidebar grouping: My Pending, My Decided (approved/rejected by me), Other Pending ──
  var pendingForMe = [];
  var decidedByMe = [];
  var otherPending = [];

  // Sort all deviations by submittedAt DESC (most recent first) before grouping
  var sortedDevs = deviations.slice().sort(function(a, b) {
    var aT = a.submittedAt ? new Date(a.submittedAt).getTime() : 0;
    var bT = b.submittedAt ? new Date(b.submittedAt).getTime() : 0;
    return bT - aT;  // desc
  });

  for (var si = 0; si < sortedDevs.length; si++) {
    var sd = sortedDevs[si];
    var myDecision = null;
    if (sd.approvals && sd.approvals.length) {
      for (var sj = 0; sj < sd.approvals.length; sj++) {
        if (String(sd.approvals[sj].approverId || '').toLowerCase() === aprIdLc) {
          myDecision = sd.approvals[sj];
          break;
        }
      }
    }
    if (myDecision) {
      decidedByMe.push({ dev: sd, decision: myDecision });
    } else if (sd.status === 'pending' || sd.status === 'partial') {
      pendingForMe.push(sd);
    } else {
      // Already approved/rejected but I never decided (rare)
      otherPending.push(sd);
    }
  }

  function buildSidebarItem(d, decisionLabel, decisionClass, isPending) {
    var devId = String(d.id || '');
    var label = decisionLabel || ('&#x23F3; Pending');
    var statusClass = decisionClass || 'sb-status-pending';
    var meta = (d.workCenter ? esc(d.workCenter) + ' &middot; ' : '') + (d.initiator ? esc(d.initiator) : '');
    return '<a href="#dev_' + esc(devId) + '" id="sb_' + esc(devId) + '" class="sb-item ' + statusClass + '">' +
      '<div class="sb-num">' + esc(d.devNum || devId) + '</div>' +
      '<div class="sb-meta">' + meta + '</div>' +
      '<div class="sb-meta" style="margin-top:3px;font-weight:700">' + label + '</div>' +
    '</a>';
  }

  var sidebarHtml = '';
  
  // Section 1: Pending for me (action required)
  sidebarHtml += '<div class="sidebar-section" id="sbAwaitSection">';
  sidebarHtml += '<div class="sidebar-section-title" id="sbAwaitTitle">&#x23F3; Awaiting Your Decision (' + pendingForMe.length + ')</div>';
  sidebarHtml += '<div id="sbAwaiting">';
  if (pendingForMe.length === 0) {
    sidebarHtml += '<div class="sb-empty" id="sbAwaitEmpty">No deviations awaiting you</div>';
  } else {
    sidebarHtml += '<div class="sb-empty" id="sbAwaitEmpty" style="display:none">No deviations awaiting you</div>';
    for (var pi = 0; pi < pendingForMe.length; pi++) {
      sidebarHtml += buildSidebarItem(pendingForMe[pi], '&#x23F3; Pending Your Decision', 'sb-status-pending', true);
    }
  }
  sidebarHtml += '</div></div>';

  // Section 2: Decided by me
  sidebarHtml += '<div class="sidebar-section" id="sbDecidedSection">';
  sidebarHtml += '<div class="sidebar-section-title" id="sbDecTitle">&#x2713; Already Decided (' + decidedByMe.length + ')</div>';
  sidebarHtml += '<div id="sbDecided">';
  if (decidedByMe.length === 0) {
    sidebarHtml += '<div class="sb-empty" id="sbDecidedEmpty">You have not decided any yet</div>';
  } else {
    sidebarHtml += '<div class="sb-empty" id="sbDecidedEmpty" style="display:none">You have not decided any yet</div>';
    for (var di = 0; di < decidedByMe.length; di++) {
      var item = decidedByMe[di];
      var lbl = item.decision.decision === 'approved'
        ? '&#x2705; You approved &middot; ' + esc(item.decision.date || '')
        : '&#x274C; You rejected &middot; ' + esc(item.decision.date || '');
      var cls = item.decision.decision === 'approved' ? 'sb-status-approved' : 'sb-status-rejected';
      sidebarHtml += buildSidebarItem(item.dev, lbl, cls, false);
    }
  }
  sidebarHtml += '</div></div>';

  // Section 3: Other (closed/decided without my involvement)
  if (otherPending.length > 0) {
    sidebarHtml += '<div class="sidebar-section">';
    sidebarHtml += '<div class="sidebar-section-title">Closed Without You (' + otherPending.length + ')</div>';
    for (var oi = 0; oi < otherPending.length; oi++) {
      var od = otherPending[oi];
      var olbl = od.status === 'approved' ? '&#x2705; Approved' : (od.status === 'rejected' ? '&#x274C; Rejected' : esc(od.status));
      var ocls = od.status === 'approved' ? 'sb-status-approved' : 'sb-status-rejected';
      sidebarHtml += buildSidebarItem(od, olbl, ocls, false);
    }
    sidebarHtml += '</div>';
  }

  // Main panel: show pending first, then decided (decided will be hidden by default)
  // Sidebar click can show them by adding 'show' class to the target card
  var orderedDevs = pendingForMe.concat(decidedByMe.map(function(x){return x.dev;}));
  // Mark which are pending vs decided via a flag we'll use when rendering cards
  var decidedIds = {};
  decidedByMe.forEach(function(x){ decidedIds[String(x.dev.id)] = true; });
  deviations = orderedDevs;

  // Build deviation cards HTML
  var cardsHtml = '';
  if (!deviations || !deviations.length) {
    cardsHtml = '<div class="empty"><div class="empty-icon">&#x1F389;</div>' +
      '<div class="empty-title">No Pending Approvals</div>' +
      '<div class="empty-sub">You have no deviations waiting for your signature right now.</div></div>';
  } else {
    for (var i = 0; i < deviations.length; i++) {
      var dev = deviations[i];
      var devId = String(dev.id || '');
      var riskColor = dev.riskFactor === 'High' ? '#c0392b' : dev.riskFactor === 'Med' ? '#c97c00' : '#2a9438';
      var alreadyApproved = false;
      var approvalRows = '';
      // Build a lookup of current approvers by id so we can show real names
      // even for old approvals that didn't store approverName.
      var allAprForLookup = getApprovers();
      var aprByIdLookup = {};
      for (var lk = 0; lk < allAprForLookup.length; lk++) {
        aprByIdLookup[String(allAprForLookup[lk].id || '').toLowerCase()] = allAprForLookup[lk];
      }
      var approvedIdsForBadge = [];
      if (dev.approvals && dev.approvals.length) {
        for (var a = 0; a < dev.approvals.length; a++) {
          var ap = dev.approvals[a];
          if (ap.decision === 'approved') approvedIdsForBadge.push(String(ap.approverId || '').toLowerCase());
          if (String(ap.approverId).toLowerCase() === String(approver.id).toLowerCase()) alreadyApproved = true;
          // Resolve display name: prefer stored approverName, else lookup by id, else email, else id
          var displayName = ap.approverName;
          if (!displayName || /^aprv_\d+$/i.test(String(displayName))) {
            var lookedUp = aprByIdLookup[String(ap.approverId || '').toLowerCase()];
            if (lookedUp && lookedUp.name) displayName = lookedUp.name;
            else if (lookedUp && lookedUp.email) displayName = lookedUp.email;
            else displayName = ap.approverId;
          }
          approvalRows += '<tr><td>' + esc(displayName) + '</td>' +
            '<td>' + (ap.decision === 'approved' ? '&#x2705; Approved' : '&#x274C; Rejected') + '</td>' +
            '<td>' + esc(ap.date) + '</td></tr>';
        }
      }
      if (!approvalRows) approvalRows = '<tr><td colspan="3" style="color:#7a8f5c;font-style:italic">No approvals yet</td></tr>';

      // Compute required count for this specific deviation (selectedApprovers + defaults)
      var devSelected = dev.selectedApprovers;
      if (typeof devSelected === 'string') {
        try { devSelected = JSON.parse(devSelected); } catch(e){ devSelected = null; }
      }
      if (!Array.isArray(devSelected) || !devSelected.length) {
        var allApr = getApprovers();
        devSelected = [];
        for (var k = 0; k < allApr.length; k++) {
          devSelected.push({
            id: allApr[k].id,
            checked: true,
            status: allApr[k].defaultStatus || (allApr[k].required ? 'required' : 'optional')
          });
        }
      }
      devSelected = devSelected.filter(function(s){ return s && s.checked !== false; });
      var requiredList = devSelected.filter(function(s){ return s.status === 'required'; });
      var ruleoutList = devSelected.filter(function(s){ return s.status === 'ruleout'; });
      var requiredCount = requiredList.length;
      var hasRuleOut = ruleoutList.length > 0;
      var totalNeeded;
      var approvedCount;
      if (hasRuleOut) {
        totalNeeded = 1;
        // Count approvals from required+ruleout (any of them satisfies the goal); cap at 1
        var eligibleIds = requiredList.concat(ruleoutList).map(function(s){return String(s.id||'').toLowerCase();});
        approvedCount = approvedIdsForBadge.filter(function(id){return eligibleIds.indexOf(id) !== -1;}).length;
        if (approvedCount > 1) approvedCount = 1;
      } else if (requiredCount > 0) {
        totalNeeded = requiredCount;
        // Only count approvals from REQUIRED approvers (optionals don't count)
        var requiredIds = requiredList.map(function(s){return String(s.id||'').toLowerCase();});
        approvedCount = approvedIdsForBadge.filter(function(id){return requiredIds.indexOf(id) !== -1;}).length;
      } else {
        totalNeeded = 1;
        approvedCount = approvedIdsForBadge.length > 0 ? 1 : 0;
      }

      var partsRows = '';
      if (dev.parts && dev.parts.length) {
        for (var p = 0; p < dev.parts.length; p++) {
          var pt = dev.parts[p];
          if (pt.partNum || pt.partName) {
            partsRows += '<tr><td>' + esc(pt.partNum) + '</td><td>' + esc(pt.qty) + '</td><td>' + esc(pt.partName) + '</td></tr>';
          }
        }
      }
      var partsHtml = partsRows
        ? '<table class="tbl"><thead><tr><th>Part #</th><th>Qty</th><th>Part Name</th></tr></thead><tbody>' + partsRows + '</tbody></table>'
        : '<p style="color:#7a8f5c;font-style:italic">None specified</p>';

      var statusBadge = dev.status === 'partial'
        ? '<span class="badge bp">&#x23F3; ' + approvedCount + '/' + totalNeeded + '</span>'
        : (hasRuleOut
            ? '<span class="badge bpend">&#x23F3; Pending (Any Rule-Out approves)</span>'
            : '<span class="badge bpend">&#x23F3; Pending (' + totalNeeded + ' needed)</span>');

      // If deviation is fully decided (approved/rejected globally), show outcome instead of form
      var devClosed = (dev.status === 'approved' || dev.status === 'rejected');
      var closedBanner = '';
      if (devClosed && !alreadyApproved) {
        var bnLabel = dev.status === 'approved' ? '&#x2705; This deviation has been approved.' : '&#x274C; This deviation was rejected.';
        var bnColor = dev.status === 'approved' ? '#2a9438' : '#c0392b';
        closedBanner = '<div class="az"><strong style="color:' + bnColor + '">' + bnLabel + ' Your decision is no longer required.</strong></div>';
      }
      var actionZone = alreadyApproved
        ? '<div class="az"><strong style="color:#2a9438">&#x2705; You already submitted your decision.</strong></div>'
        : '<div class="az">' +
            '<div style="font-size:16px;font-weight:700;color:#7A9A3A;margin-bottom:12px">&#x1F4DD; Your Decision</div>' +
            '<div id="sb_' + devId + '" style="display:none;background:#39B54A;color:#fff;padding:14px;border-radius:8px;font-weight:700;margin-bottom:10px"></div>' +
            '<div id="fm_' + devId + '">' +
              '<div class="fld"><label>Decision</label>' +
                '<select id="dc_' + devId + '">' +
                  '<option value="approved">&#x2705; Approve</option>' +
                  '<option value="rejected">&#x274C; Reject</option>' +
                '</select></div>' +
              '<div class="fld"><label>Comments (optional)</label>' +
                '<textarea id="cm_' + devId + '" placeholder="Add notes..."></textarea></div>' +
              '<button class="btn" id="bs_' + devId + '" onclick="sub(\x27' + devId + '\x27)">Submit Decision</button>' +
            '</div>' +
          '</div>';

      var isDecidedCard = decidedIds[String(devId)] === true;
      var cardClass = isDecidedCard ? 'card card-decided' : 'card';
      cardsHtml +=
        '<div class="' + cardClass + '" id="dev_' + esc(devId) + '">' +
          '<div class="card-hdr">' +
            '<div><div class="devnum">' + esc(dev.devNum) + '</div>' +
            '<div style="font-size:13px;color:#7a8f5c">WC: ' + esc(dev.workCenter) + ' &middot; ' + esc(dev.initiator) + '</div></div>' +
            statusBadge +
          '</div>' +
          '<div class="grid5">' +
            '<div><div class="lbl">Date</div><div class="val">' + esc(fmtUS(dev.date)) + '</div></div>' +
            '<div><div class="lbl">Start</div><div class="val">' + esc(fmtUS(dev.startDate)) + '</div></div>' +
            '<div><div class="lbl">End</div><div class="val">' + esc(fmtUS(dev.endDate)) + '</div></div>' +
            '<div><div class="lbl">Risk</div><div class="val" style="color:' + riskColor + ';font-weight:700">' + esc(dev.riskFactor || 'Low') + '</div></div>' +
            '<div><div class="lbl">Cust.Approval</div><div class="val">' + esc(dev.custApproval || 'No') + '</div></div>' +
          '</div>' +
          '<div class="sec"><div class="lbl">Description</div><div class="txt">' + esc(dev.description) + '</div></div>' +
          '<div class="sec"><div class="lbl">Parts</div>' + partsHtml + '</div>' +
          '<div class="sec"><div class="lbl">Reason</div><div>' + esc(dev.reasonLabel) + '</div></div>' +
          '<div class="sec"><div class="lbl">Action Plan</div><div class="txt">' + esc(dev.actionPlan) + '</div>' +
            (dev.comments ? '<div style="margin-top:8px;padding-top:8px;border-top:1px dashed #d4e0b8"><div class="lbl" style="margin-bottom:4px">Comments / Special Instructions</div><div class="txt">' + esc(dev.comments) + '</div></div>' : '') +
            (dev.owner ? '<div style="margin-top:8px;padding-top:8px;border-top:1px dashed #d4e0b8"><div class="lbl" style="margin-bottom:4px">Owner / Responsible Party</div><div class="txt">' + esc(dev.owner) + '</div></div>' : '') +
          '</div>' +
          (function(){
            var photos = dev.photos;
            if (typeof photos === 'string') {
              try { photos = JSON.parse(photos); } catch(e) { photos = []; }
            }
            if (!Array.isArray(photos)) photos = [];
            photos = photos.filter(function(p) { return p && p.dataUrl; });
            if (!photos.length) return '';
            // Use unique key per-deviation to avoid conflicts in the same page
            var devKey = String(dev.id || '').replace(/[^a-zA-Z0-9_]/g, '_');
            var photoHtml = '<div class="sec"><div class="lbl">Photo Evidence (' + photos.length + ')</div>' +
              '<div class="photos-grid">';
            for (var pi = 0; pi < photos.length; pi++) {
              var p = photos[pi];
              // The image src has the dataUrl already; click handler reads from the DOM (avoids inline-string issues with apostrophes/quotes in base64)
              photoHtml += '<div class="photo-cell" onclick="showLightbox(this)" title="Click to enlarge">' +
                '<img src="' + p.dataUrl + '" alt="evidence ' + (pi+1) + '">' +
              '</div>';
            }
            photoHtml += '</div></div>';
            return photoHtml;
          })() +
          '<div class="sec"><div class="lbl">Approvals (' + (dev.approvals || []).length + ')</div>' +
            '<table class="tbl"><thead><tr><th>Approver</th><th>Decision</th><th>Date</th></tr></thead>' +
            '<tbody>' + approvalRows + '</tbody></table>' +
          '</div>' +
          (devClosed && !alreadyApproved ? closedBanner : actionZone) +
        '</div>';
    }
  }

  var initials = '';
  var nameParts = String(approver.name || '').split(' ');
  for (var n = 0; n < nameParts.length && n < 2; n++) initials += nameParts[n][0] || '';
  initials = initials.toUpperCase();

  var approverJson = JSON.stringify(approver);
  var tokenJson    = JSON.stringify(token);
  var countText    = deviations.length + ' deviation' + (deviations.length !== 1 ? 's' : '') + ' awaiting your signature';

  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Pending Approvals</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Calibri,Arial,sans-serif;background:#f4f7f0;color:#2c3e1a;font-size:15px}' +
    '.hdr{background:#fff;border-bottom:3px solid #7A9A3A;padding:0 24px;height:60px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(90,120,40,.1)}' +
    '.logo{font-size:18px;font-weight:700;color:#7A9A3A}' +
    '.chip{background:#e6f6f8;border:1.5px solid #007b8a;border-radius:8px;padding:6px 14px;font-size:13px;font-weight:600;color:#005e6b}' +
    '.layout{display:flex;gap:20px;max-width:1280px;margin:0 auto;padding:20px;align-items:flex-start}' +
    '.sidebar{flex:0 0 280px;background:#fff;border:1.5px solid #d4e0b8;border-radius:12px;padding:14px;box-shadow:0 2px 8px rgba(90,120,40,.06);position:sticky;top:80px;max-height:calc(100vh - 100px);overflow-y:auto}' +
    '.sidebar-title{font-size:13px;font-weight:700;color:#7A9A3A;text-transform:uppercase;margin-bottom:10px;padding-bottom:8px;border-bottom:1.5px solid #d4e0b8}' +
    '.sb-search-wrap{position:relative;margin-bottom:12px}' +
    '.sb-search{width:100%;padding:7px 28px 7px 10px;border:1.5px solid #d4e0b8;border-radius:6px;font-size:12px;font-family:inherit;background:#f8faf4;color:#2c3e1a;outline:none;box-sizing:border-box}' +
    '.sb-search:focus{border-color:#7A9A3A;background:#fff}' +
    '.sb-search-clear{position:absolute;right:6px;top:50%;transform:translateY(-50%);background:none;border:none;color:#7a8f5c;cursor:pointer;font-size:13px;padding:2px 6px;line-height:1}' +
    '.sb-search-clear:hover{color:#c0392b}' +
    '.sb-item.sb-hidden{display:none}' +
    '.sidebar-section.sb-section-empty{display:none}' +
    '.sidebar-section{margin-bottom:14px}' +
    '.sidebar-section-title{font-size:10px;font-weight:700;color:#7a8f5c;text-transform:uppercase;margin-bottom:6px}' +
    '.sb-item{display:block;padding:8px 10px;border-radius:6px;cursor:pointer;font-size:12px;color:#2c3e1a;border:1.5px solid transparent;margin-bottom:4px;text-decoration:none;transition:all 0.15s}' +
    '.sb-item:hover{background:#f2f7e8;border-color:#d4e0b8}' +
    '.sb-item.sb-active{background:#7A9A3A;color:#fff;border-color:#7A9A3A}' +
    '.sb-item .sb-num{font-weight:700;font-size:13px}' +
    '.sb-item .sb-meta{font-size:10px;opacity:0.8;margin-top:2px}' +
    '.sb-empty{font-size:11px;color:#7a8f5c;font-style:italic;padding:6px}' +
    '.sb-status-pending{border-left:3px solid #c97c00}' +
    '.sb-status-partial{border-left:3px solid #c97c00}' +
    '.sb-status-approved{border-left:3px solid #2a9438}' +
    '.sb-status-rejected{border-left:3px solid #c0392b}' +
    '.main{flex:1;min-width:0;padding:0}' +
    '@media(max-width:840px){.layout{flex-direction:column;padding:14px}.sidebar{position:static;flex:0;width:100%;max-height:none}.main{padding:0}}' +
    '.title{font-size:24px;font-weight:700;color:#7A9A3A;margin-bottom:4px}' +
    '.sub{font-size:13px;color:#7a8f5c;margin-bottom:20px}' +
    '.rolebox{background:#e6f6f8;border:1.5px solid #007b8a;border-radius:8px;padding:12px 16px;margin-bottom:20px;font-size:13px;color:#005e6b}' +
    '.card{background:#fff;border:1.5px solid #d4e0b8;border-radius:12px;margin-bottom:20px;overflow:hidden;box-shadow:0 2px 8px rgba(90,120,40,.08);scroll-margin-top:80px}' +
    '.card-decided{display:none}' +
    '.card-decided:target{display:block;border-color:#7a8f5c;background:#fafbf6}' +
    '.card-decided.show{display:block;border-color:#7a8f5c;background:#fafbf6}' +
    '.card-hdr{background:#f2f7e8;border-bottom:1.5px solid #d4e0b8;padding:14px 18px;display:flex;justify-content:space-between;align-items:flex-start}' +
    '.devnum{font-size:20px;font-weight:700;color:#7A9A3A}' +
    '.badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700}' +
    '.bpend{background:#fff8e1;color:#c97c00;border:1px solid #f0c050}' +
    '.bp{background:#e6f6f8;color:#005e6b;border:1px solid #007b8a}' +
    '.grid5{display:grid;grid-template-columns:repeat(3,1fr);border-bottom:1.5px solid #d4e0b8}' +
    '.grid5>div{padding:10px 14px;border-right:1px solid #d4e0b8}' +
    '.grid5>div:last-child{border-right:none}' +
    '.lbl{font-size:10px;font-weight:700;color:#7A9A3A;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}' +
    '.val{font-size:14px;font-weight:600}' +
    '.sec{padding:14px 18px;border-bottom:1px solid #d4e0b8}' +
    '.photos-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;margin-top:6px}' +
    '.photo-cell{background:#f4f7f0;border:1.5px solid #d4e0b8;border-radius:6px;overflow:hidden;cursor:pointer;display:flex;align-items:center;justify-content:center;min-height:120px;max-height:180px;transition:all 0.2s}' +
    '.photo-cell:hover{border-color:#7A9A3A;transform:scale(1.02)}' +
    '.photo-cell img{max-width:100%;max-height:180px;object-fit:contain;display:block}' +
    '.txt{background:#f2f7e8;border-left:3px solid #7A9A3A;padding:8px 12px;border-radius:0 6px 6px 0;font-size:14px;line-height:1.5}' +
    '.tbl{width:100%;border-collapse:collapse;font-size:13px}' +
    '.tbl th{background:#7A9A3A;color:#fff;padding:7px 12px;text-align:left;font-size:11px;text-transform:uppercase}' +
    '.tbl td{padding:7px 12px;border-bottom:1px solid #d4e0b8}' +
    '.az{padding:18px;background:#edfbef;border-top:2px solid #39B54A}' +
    '.fld{margin-bottom:12px}' +
    '.fld label{display:block;font-size:11px;font-weight:700;color:#7A9A3A;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}' +
    '.fld select,.fld textarea{width:100%;padding:9px 12px;border:1.5px solid #d4e0b8;border-radius:8px;font-family:Calibri,Arial,sans-serif;font-size:14px;background:#fff}' +
    'textarea{resize:vertical;min-height:70px}' +
    '.btn{background:#39B54A;color:#fff;border:none;border-radius:8px;padding:11px 22px;font-size:14px;font-weight:700;cursor:pointer;min-height:44px}' +
    '.btn:disabled{opacity:.5;cursor:not-allowed}' +
    '.empty{background:#fff;border:1.5px solid #d4e0b8;border-radius:12px;padding:48px;text-align:center}' +
    '.empty-icon{font-size:48px;margin-bottom:12px}' +
    '.empty-title{font-size:20px;font-weight:700;color:#7A9A3A;margin-bottom:8px}' +
    '.empty-sub{color:#7a8f5c;font-size:14px}' +
    '.tc{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);z-index:500;display:flex;flex-direction:column;gap:8px;min-width:260px}' +
    '.toast{padding:12px 16px;border-radius:8px;font-size:14px;font-weight:600;background:#fff;border:1.5px solid #d4e0b8}' +
    '.ts{background:#edfbef;border-color:rgba(57,181,74,.4)}' +
    '.te{background:#fdf0ee;border-color:rgba(192,57,43,.3)}' +
    '</style></head><body>' +
    '<div class="hdr"><div class="logo">MWAAF &mdash; Deviation Authorization</div>' +
    '<div class="chip">' + initials + ' &nbsp; ' + esc(approver.name) + '</div></div>' +
    '<div class="layout">' +
    '<div class="sidebar">' +
    '<div class="sidebar-title">&#x1F4CB; All Deviations</div>' +
    '<div class="sb-search-wrap">' +
      '<input type="text" id="sbSearch" class="sb-search" placeholder="Search by Dev #, Part #, Work Center..." autocomplete="off" oninput="filterSidebar(this.value)">' +
      '<button type="button" class="sb-search-clear" onclick="document.getElementById(\'sbSearch\').value=\'\';filterSidebar(\'\');" title="Clear">&#x2715;</button>' +
    '</div>' +
    sidebarHtml +
    '</div>' +
    '<div class="main">' +
    '<div class="title">Your Approvals</div>' +
    '<div class="sub">' + (deviations.length ? countText : 'All caught up!') + '</div>' +
    '<div class="rolebox">Signed in as: <strong>' + esc(approver.name) + '</strong> &nbsp;|&nbsp; Role: <strong>' + esc(approver.role) + '</strong></div>' +
    cardsHtml +
    '</div>' +
    '</div>' +
    '<div class="tc" id="tc"></div>' +
    /* Lightbox modal for image preview */
    '<div id="imgLightbox" onclick="hideLightbox()" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.92);z-index:9999;align-items:center;justify-content:center;padding:20px;cursor:zoom-out">' +
      '<div style="position:relative;max-width:96vw;max-height:96vh;display:flex;flex-direction:column;align-items:center" onclick="event.stopPropagation()">' +
        '<button onclick="hideLightbox()" style="position:absolute;top:-40px;right:0;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);color:#fff;border-radius:8px;padding:6px 14px;cursor:pointer;font-size:14px;font-weight:600">&#x2715; Close</button>' +
        '<img id="imgLightboxImg" src="" style="max-width:96vw;max-height:90vh;object-fit:contain;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,0.5)">' +
      '</div>' +
    '</div>' +
    '<script>' +
    /* Lightbox functions */
    'function showLightbox(cell){' +
      'var img = cell.querySelector("img");' +
      'if (!img) return;' +
      'var lb = document.getElementById("imgLightbox");' +
      'var lbImg = document.getElementById("imgLightboxImg");' +
      'lbImg.src = img.src;' +
      'lb.style.display = "flex";' +
      'document.body.style.overflow = "hidden";' +
    '}' +
    'function hideLightbox(){' +
      'var lb = document.getElementById("imgLightbox");' +
      'if (lb) lb.style.display = "none";' +
      'document.body.style.overflow = "";' +
    '}' +
    'document.addEventListener("keydown", function(e){ if (e.key === "Escape") hideLightbox(); });' +
    'var APPROVER=' + approverJson + ';' +
    'var TOKEN=' + tokenJson + ';' +
    'var submitted={};' +
    'function sub(id){' +
      'if(submitted[id]){toast("Already submitted.","");return;}' +
      'var d=document.getElementById("dc_"+id).value;' +
      'var c=(document.getElementById("cm_"+id).value||"").trim();' +
      'var b=document.getElementById("bs_"+id);' +
      'b.disabled=true;b.textContent="Submitting...";' +
      'google.script.run' +
        '.withSuccessHandler(function(result){' +
          'submitted[id]=true;' +
          'updateUiAfterDecision(id, d, c, result);' +
          /* Show server-side info message if present, else default */
          'var infoMsg = (result && result.info) ? result.info : (d==="approved"?"\u2705 Approval recorded!":"\u274C Rejection recorded.");' +
          'var toastClass = (result && result.statusAfter==="approved") ? "ts" : (d==="approved"?"ts":"te");' +
          'toast(infoMsg, toastClass);' +
        '})' +
        '.withFailureHandler(function(e){' +
          'b.disabled=false;b.textContent="Submit Decision";' +
          'toast("Error: "+(e.message||String(e)),"te");' +
        '})' +
        '.submitApprovalFromView(id,APPROVER.id,d,c,TOKEN);' +
    '}' +
    /* Transform card to "decided" state + move sidebar item */
    'function updateUiAfterDecision(id, decision, comments, result){' +
      'var card = document.getElementById("dev_"+id);' +
      'var sbItem = document.getElementById("sb_"+id);' +
      'var dateStr = new Date().toLocaleDateString();' +
      'var statusAfter = (result && result.statusAfter) || "";' +
      'var infoMsg = (result && result.info) || "";' +
      /* Determine display label and color based on FINAL deviation status, not just my decision */
      'var lbl, clsNew, bannerStyle, bannerText;' +
      'if (decision === "approved") {' +
        'lbl = "\u2705 You approved \u00B7 "+dateStr;' +
        'clsNew = "sb-status-approved";' +
        'bannerStyle = "background:#edfbef;border-color:#39B54A;color:#2a9438";' +
        'bannerText = "\u2705 You approved this deviation on " + dateStr;' +
        'if (statusAfter === "approved" && infoMsg && infoMsg.indexOf("override") !== -1) {' +
          /* Ruleout overrode rejection */
          'bannerText += " \u2014 Your rule-out approval overrode a prior rejection. Deviation now APPROVED.";' +
        '}' +
      '} else {' +
        'lbl = "\u274C You rejected \u00B7 "+dateStr;' +
        'clsNew = "sb-status-rejected";' +
        /* If statusAfter is approved (deviation was already approved), show a softer message */
        'if (statusAfter === "approved") {' +
          'bannerStyle = "background:#fff8e1;border-color:#c97c00;color:#7a5000";' +
          'bannerText = "\u26A0 You rejected on " + dateStr + ", but this deviation was already approved by required approvers \u2014 your rejection is logged but does not change the status.";' +
        '} else if (infoMsg && infoMsg.indexOf("optional rejection") !== -1) {' +
          /* Optional rejection that doesn\'t block */
          'bannerStyle = "background:#fff8e1;border-color:#c97c00;color:#7a5000";' +
          'bannerText = "\u26A0 You rejected on " + dateStr + ". Optional rejections are logged but do not block approval \u2014 the deviation continues to need required signatures.";' +
        '} else {' +
          'bannerStyle = "background:#fdf0ee;border-color:#c0392b;color:#c0392b";' +
          'bannerText = "\u274C You rejected this deviation on " + dateStr;' +
        '}' +
      '}' +
      /* 1. Move sidebar item from "Awaiting" to "Already Decided" section */
      'if (sbItem) {' +
        'sbItem.parentNode.removeChild(sbItem);' +
        'var decidedSection = document.getElementById("sbDecided");' +
        'var decidedEmpty = document.getElementById("sbDecidedEmpty");' +
        'if (decidedEmpty) decidedEmpty.style.display="none";' +
        'if (decidedSection) {' +
          'sbItem.className = "sb-item "+clsNew;' +
          'var lblDiv = sbItem.querySelector(".sb-meta:last-child");' +
          'if (lblDiv) lblDiv.innerHTML = lbl;' +
          'decidedSection.insertBefore(sbItem, decidedSection.firstChild);' +
        '}' +
        /* Update counters */
        'var awaitTitle = document.getElementById("sbAwaitTitle");' +
        'var decTitle = document.getElementById("sbDecTitle");' +
        'if (awaitTitle) {' +
          'var cAwait = document.querySelectorAll("#sbAwaiting .sb-item").length;' +
          'awaitTitle.innerHTML = "\u23F3 Awaiting Your Decision ("+cAwait+")";' +
          'var awEmpty = document.getElementById("sbAwaitEmpty");' +
          'if (awEmpty) awEmpty.style.display = cAwait===0 ? "block" : "none";' +
        '}' +
        'if (decTitle) {' +
          'var cDec = document.querySelectorAll("#sbDecided .sb-item").length;' +
          'decTitle.innerHTML = "\u2713 Already Decided ("+cDec+")";' +
        '}' +
      '}' +
      /* 2. Transform the card into a "decided" state (keep visible) */
      'if (card) {' +
        'var formEl = document.getElementById("fm_"+id);' +
        'if (formEl) formEl.style.display = "none";' +
        'var banner = document.createElement("div");' +
        'banner.className = "decision-banner";' +
        'banner.style.cssText = "margin:0 0 12px 0;padding:10px 14px;border-radius:8px;font-weight:700;font-size:13px;border:2px solid;" + bannerStyle;' +
        'banner.innerHTML = bannerText + (comments ? "<br><span style=\\"font-weight:400;font-size:12px;color:#555\\">Your comments: " + comments.replace(/</g, "&lt;") + "</span>" : "");' +
        'card.insertBefore(banner, card.firstChild);' +
        'card.style.opacity = "0.92";' +
      '}' +
      /* 3. Update main panel header count */
      'var subEl = document.querySelector(".sub");' +
      'var visibleForms = 0;' +
      'var allForms = document.querySelectorAll("[id^=\\"fm_\\"]");' +
      'for (var fi=0; fi<allForms.length; fi++) {' +
        'if (allForms[fi].style.display !== "none") visibleForms++;' +
      '}' +
      'if (subEl) {' +
        'if (visibleForms <= 0) subEl.innerHTML = "\u2705 All caught up! No deviations awaiting your decision.";' +
        'else subEl.innerHTML = visibleForms + " deviation"+(visibleForms!==1?"s":"")+" awaiting your decision";' +
      '}' +
    '}' +
    'function toast(msg,cls){' +
      'var c=document.getElementById("tc");' +
      'var t=document.createElement("div");' +
      't.className="toast "+(cls||"");' +
      't.textContent=msg;' +
      't.onclick=function(){t.parentNode&&t.parentNode.removeChild(t);};' +
      'c.appendChild(t);' +
      'setTimeout(function(){if(t.parentNode)t.parentNode.removeChild(t);},5000);' +
    '}' +
    /* Sidebar live search — matches Dev #, Part #, Work Center, Initiator */
    'function filterSidebar(q){' +
      'q = (q||"").trim().toLowerCase();' +
      'var items = document.querySelectorAll(".sidebar .sb-item");' +
      'for (var i=0;i<items.length;i++){' +
        'var it = items[i];' +
        'var text = (it.textContent||"").toLowerCase();' +
        'if (!q || text.indexOf(q) !== -1) it.classList.remove("sb-hidden");' +
        'else it.classList.add("sb-hidden");' +
      '}' +
      /* Hide empty sections */
      'var sections = document.querySelectorAll(".sidebar .sidebar-section");' +
      'for (var j=0;j<sections.length;j++){' +
        'var sec = sections[j];' +
        'var visible = sec.querySelectorAll(".sb-item:not(.sb-hidden)").length;' +
        'if (q && visible === 0) sec.classList.add("sb-section-empty");' +
        'else sec.classList.remove("sb-section-empty");' +
      '}' +
    '}' +
    '</script></body></html>';
}

// ── SETUP / INIT ───────────────────────────────────────────

/**
 * Run this function ONCE from the Apps Script editor to:
 *   1. Create all required sheets with headers
 *   2. Seed default approvers, config, and reason options
 * Menu: Run → setupSheets
 */

/** Email sender wrapper — uses MailApp by default; tries Gmail alias only when FROM_EMAIL is set */
function _sendEmailSafe(to, subject, htmlBody, cc) {
  // If FROM_EMAIL is configured AND verified as alias, try GmailApp with from:
  if (FROM_EMAIL) {
    try {
      const opts = {
        htmlBody: htmlBody,
        name: FROM_NAME,
        from: FROM_EMAIL
      };
      if (cc) opts.cc = cc;
      GmailApp.sendEmail(to, subject, '', opts);
      Logger.log('Email sent via GmailApp(from=' + FROM_EMAIL + ') to: ' + to);
      return;
    } catch (e) {
      Logger.log('GmailApp w/ alias failed (' + e.message + '), falling back to MailApp');
    }
  }
  
  // Default path: MailApp (sends from script owner with friendly name)
  const mailOpts = {
    to: to,
    subject: subject,
    htmlBody: htmlBody,
    name: FROM_NAME
  };
  if (cc) mailOpts.cc = cc;
  MailApp.sendEmail(mailOpts);
  Logger.log('Email sent via MailApp to: ' + to);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Save sheet ID immediately so Web App can find it later
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('');
  Logger.log('IMPORTANT: After deploying as Web App, run setWebAppUrl() to save the URL.');
  Logger.log('Approval email links will not work until you do this.');

  // Helper to create sheet if missing
  function ensureSheet(name, headers, rows) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clearContents();
    const allRows = rows ? [headers, ...rows] : [headers];
    sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    Logger.log('Sheet created/reset: ' + name);
    return sheet;
  }

  // Deviations
  ensureSheet(SH.DEVIATIONS, DEV_HEADERS);

  // Approvals
  ensureSheet(SH.APPROVALS, APPR_HEADERS);

  // Tokens — force plain text format to prevent date auto-conversion
  const tokSheet = ensureSheet(SH.TOKENS, ['token','approverId','approverEmail','createdAt','expiresAt','used']);
  if (tokSheet) {
    tokSheet.getRange(1, 1, tokSheet.getMaxRows(), 6).setNumberFormat('@STRING@');
  }

  // Approvers (default 5 roles) — use 'aprv' prefix to avoid Sheets confusing 'apr' with April
  // 6 columns: id, name, email, role, required, defaultStatus
  ensureSheet(SH.APPROVERS, APR_HEADERS, [
    ['aprv1', 'Omar Segura',       'omar@mwaaf.com',        "Quality Manager or assigned TM's", true,  'required'],
    ['aprv2', 'Heather Teters',    'heather@mwaaf.com',     "Production or assigned TM's",       true,  'required'],
    ['aprv3', 'Engineering Lead',  'engineering@mwaaf.com', 'Engineering/Processing',            false, 'optional'],
    ['aprv4', 'Materials Manager', 'materials@mwaaf.com',   'Materials',                         false, 'optional'],
    ['aprv5', 'Plant Manager',     'plantmgr@mwaaf.com',    'Plant Manager',                     false, 'optional'],
  ]);

  // Config
  ensureSheet(SH.CONFIG, ['key','value'], [
    ['docOwner',          'Director of Quality'],
    ['approvedByTitle',   'Director of Operations'],
    ['draftValidHours',   '12'],
    ['settingsPassword',  'mwaaf2024'],
    ['appAccessPassword', 'Mwaaf01'],
    ['nextDevNum',        '1001'],
  ]);

  // Role Config
  ensureSheet(SH.ROLE_CONFIG, ['roleKey','required'], [
    ["Quality Manager or assigned TM's", true],
    ["Production or assigned TM's",       true],
    ['Engineering/Processing',            false],
    ['Materials',                         false],
    ['Plant Manager',                     false],
  ]);

  // Reason Options
  ensureSheet(SH.REASON_OPTIONS, ['id','label','tags'], [
    ['r1','1. Use parts that do not meet engineering drawing specifications','[]'],
    ['r2','2. Material substitution (Describe in Comments/Special Instructions)','[]'],
    ['r3','3. Use parts that do not meet approved appearance standards','[]'],
    ['r4','4. Use parts for saleable vehicles prior to Supplier Part Submission approval','[]'],
    ['r5','5. Sample build on current production line','[]'],
    ['r6','6. Temporary Deviation from the CCS, PWI, QWI, SWI, C/O Set-up, etc.','[]'],
    ['r7','7. Continuation of non-capable process','[]'],
    ['r8','8. 4M Change Category (specify below)','["4M"]'],
    ['r9','9. Other — specify reason below','["other"]'],
  ]);

  // Dist Lists
  ensureSheet(SH.DIST_LISTS, ['listType','email'], [
    ['creation', 'ops@mwaaf.com'],
    ['creation', 'quality@mwaaf.com'],
    ['approval', 'ops@mwaaf.com'],
    ['approval', 'quality@mwaaf.com'],
    ['approval', 'production@mwaaf.com'],
  ]);

  // Part Numbers Catalog (force text format)
  const pnSheet = ensureSheet(SH.PART_NUMBERS, ['partNumber','description'], [
    ['44521', ''],
    ['12345', ''],
  ]);
  if (pnSheet) {
    pnSheet.getRange(1, 1, pnSheet.getMaxRows(), 2).setNumberFormat('@');
  }

  // Work Centers Catalog (force text format)
  const wcSheet = ensureSheet(SH.WORK_CENTERS, ['workCenter'], [
    ['2240'],
    ['2250'],
    ['2260'],
  ]);
  if (wcSheet) {
    wcSheet.getRange(1, 1, wcSheet.getMaxRows(), 1).setNumberFormat('@');
  }

  // Set column widths for Deviations sheet
  const devSheet = ss.getSheetByName(SH.DEVIATIONS);
  if (devSheet) {
    devSheet.setColumnWidth(1, 160);   // id
    devSheet.setColumnWidth(2, 90);    // devNum
    devSheet.setColumnWidth(3, 100);   // mainPartNum
    devSheet.setColumnWidth(12, 300);  // description
    devSheet.setColumnWidth(20, 300);  // actionPlan
  }


  Logger.log('=== SETUP COMPLETE ===');
  Logger.log('All sheets created successfully.');
  Logger.log('');
  Logger.log('NEXT STEPS:');
  Logger.log('1. Go to Deploy > New deployment');
  Logger.log('2. Type: Web app | Execute as: Me | Access: Anyone');
  Logger.log('3. Click Deploy and copy the URL');
  Logger.log('4. Paste that URL into MY_WEB_APP_URL at the top of Code.gs');
  Logger.log('5. Update approver emails in the Approvers sheet');
  Logger.log('6. Change settingsPassword and appAccessPassword in the Config sheet');
  Logger.log('');
  Logger.log('Sheet ID saved: ' + ss.getId());

  // Also try to show UI alert if running from spreadsheet context
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Setup Complete',
      'All sheets have been created!\n\n' +
      'Next steps:\n' +
      '1. Deploy > New deployment > Web App\n' +
      '2. Copy the URL\n' +
      '3. Paste it in MY_WEB_APP_URL at the top of Code.gs\n' +
      '4. Update emails in the Approvers sheet',
      ui.ButtonSet.OK
    );
  } catch(uiErr) {
    // Running from editor without sheet context — logs above are sufficient
  }
}

/**
 * ─── RUN THIS ONCE AFTER EACH NEW DEPLOYMENT ───
 * Paste your new Web App URL inside the quotes below, then Run this function.
 * This saves it permanently in Script Properties — no need to touch any other file.
 */


// ── REPAIR UTILITY ──────────────────────────────────────────────
// Run this manually from the Apps Script editor to fix sheet data corruption issues:
// - Approver IDs stored with wrong case (Apr1, Apr6, etc.)
// - Approver IDs auto-converted to dates (apr1 → 2026-04-01)
// - Token sheet with date-corrupted approverIds
// 
// Usage: open Apps Script editor, select 'repairSheets' from the function dropdown, click Run.
// ── ID PREFIX MIGRATION ──────────────────────────────────────
// Migrates all approver IDs from "apr" prefix to "aprv" prefix throughout ALL sheets.
// This permanently fixes the Google Sheets auto-conversion of "apr1" to a date.
// Run once from Apps Script editor: select 'migrateAprToAprv' and click Run.
function migrateAprToAprv() {
  const result = { approvers: 0, approvals: 0, tokens: 0, deviations: 0, mapping: {} };

  // Build mapping: old id (lowercase) → new id
  const aprSheet = getSheet(SH.APPROVERS, true);
  const aprData  = aprSheet.getDataRange().getValues();
  if (aprData.length < 2) {
    Logger.log('No approvers to migrate');
    return result;
  }
  
  // Force entire approvers sheet to text format
  aprSheet.getRange(1, 1, aprSheet.getMaxRows(), aprData[0].length).setNumberFormat('@');
  
  const aprHeaders = aprData[0].map(h => String(h));
  const aprIdCol   = aprHeaders.indexOf('id');
  
  // First pass: build mapping and rename in Approvers sheet
  for (let i = 1; i < aprData.length; i++) {
    const oldVal = aprData[i][aprIdCol];
    let oldId = '';
    if (oldVal instanceof Date) {
      // Date-corrupted, infer from row position
      oldId = 'apr' + i;
    } else {
      oldId = String(oldVal || '').toLowerCase().trim();
    }
    
    // Only migrate if old id starts with 'apr' but NOT already 'aprv'
    if (oldId.match(/^apr\d/) && !oldId.startsWith('aprv')) {
      const newId = oldId.replace(/^apr/, 'aprv');
      result.mapping[oldId] = newId;
      aprSheet.getRange(i + 1, aprIdCol + 1).setNumberFormat('@').setValue("'" + newId);
      result.approvers++;
      Logger.log('Approvers: ' + oldId + ' → ' + newId);
    } else if (/^\d{4}-\d{2}-\d{2}/.test(oldId) || /^\d+\/\d+\/\d{4}/.test(oldId) || oldVal instanceof Date) {
      // Date-corrupted — recreate as aprv<row>
      const newId = 'aprv' + i;
      result.mapping[oldId] = newId;
      aprSheet.getRange(i + 1, aprIdCol + 1).setNumberFormat('@').setValue("'" + newId);
      result.approvers++;
      Logger.log('Approvers (date-corrupt): ' + oldId + ' → ' + newId);
    }
  }
  
  Logger.log('ID mapping: ' + JSON.stringify(result.mapping));
  
  // Second pass: update Approvals sheet
  const apvlSheet = getSheet(SH.APPROVALS, false);
  if (apvlSheet) {
    const apvlData = apvlSheet.getDataRange().getValues();
    if (apvlData.length > 1) {
      apvlSheet.getRange(1, 1, apvlSheet.getMaxRows(), apvlData[0].length).setNumberFormat('@');
      const apvlHeaders = apvlData[0].map(h => String(h));
      const apvlAprIdCol = apvlHeaders.indexOf('approverId');
      
      if (apvlAprIdCol >= 0) {
        for (let j = 1; j < apvlData.length; j++) {
          const oldVal = apvlData[j][apvlAprIdCol];
          let oldKey = String(oldVal || '').toLowerCase().trim();
          
          // Direct mapping first
          if (result.mapping[oldKey]) {
            apvlSheet.getRange(j + 1, apvlAprIdCol + 1).setNumberFormat('@').setValue("'" + result.mapping[oldKey]);
            result.approvals++;
          } else if (oldKey.match(/^apr\d/) && !oldKey.startsWith('aprv')) {
            // Same pattern apr<n> → aprv<n>
            const newId = oldKey.replace(/^apr/, 'aprv');
            apvlSheet.getRange(j + 1, apvlAprIdCol + 1).setNumberFormat('@').setValue("'" + newId);
            result.approvals++;
          }
        }
      }
    }
  }
  
  // Third pass: update Tokens sheet
  const tokSheet = getSheet(SH.TOKENS, false);
  if (tokSheet) {
    const tokData = tokSheet.getDataRange().getValues();
    if (tokData.length > 1) {
      tokSheet.getRange(1, 1, tokSheet.getMaxRows(), tokData[0].length).setNumberFormat('@');
      const tokHeaders = tokData[0].map(h => String(h));
      const tokAprCol = tokHeaders.indexOf('approverId');
      
      if (tokAprCol >= 0) {
        for (let k = 1; k < tokData.length; k++) {
          const oldVal = tokData[k][tokAprCol];
          let oldKey = String(oldVal || '').toLowerCase().trim();
          
          if (result.mapping[oldKey]) {
            tokSheet.getRange(k + 1, tokAprCol + 1).setNumberFormat('@').setValue("'" + result.mapping[oldKey]);
            result.tokens++;
          } else if (oldKey.match(/^apr\d/) && !oldKey.startsWith('aprv')) {
            const newId = oldKey.replace(/^apr/, 'aprv');
            tokSheet.getRange(k + 1, tokAprCol + 1).setNumberFormat('@').setValue("'" + newId);
            result.tokens++;
          }
        }
      }
    }
  }
  
  // Fourth pass: update selectedApprovers JSON in Deviations sheet
  const devSheet = getSheet(SH.DEVIATIONS, false);
  if (devSheet) {
    const devData = devSheet.getDataRange().getValues();
    if (devData.length > 1) {
      const devHeaders = devData[0].map(h => String(h));
      const selCol = devHeaders.indexOf('selectedApprovers');
      
      if (selCol >= 0) {
        for (let m = 1; m < devData.length; m++) {
          const oldStr = String(devData[m][selCol] || '');
          if (!oldStr) continue;
          
          let updated = oldStr;
          let didUpdate = false;
          
          // Replace each apr<n> reference (but not aprv<n>) with aprv<n>
          // Use a regex that matches "id":"apr<digits>" but NOT "aprv<digits>"
          updated = updated.replace(/"id"\s*:\s*"apr(\d[^"]*)"/g, function(match, rest) {
            didUpdate = true;
            return '"id":"aprv' + rest + '"';
          });
          
          // Also fix any direct mapping entries
          Object.keys(result.mapping).forEach(oldKey => {
            const newKey = result.mapping[oldKey];
            // Match the id value in JSON
            const re = new RegExp('"id"\\s*:\\s*"' + oldKey.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&') + '"', 'g');
            if (re.test(updated)) {
              updated = updated.replace(re, '"id":"' + newKey + '"');
              didUpdate = true;
            }
          });
          
          if (didUpdate) {
            devSheet.getRange(m + 1, selCol + 1).setValue(updated);
            result.deviations++;
            Logger.log('Deviations row ' + (m+1) + ' selectedApprovers updated');
          }
        }
      }
    }
  }
  
  Logger.log('=== MIGRATION COMPLETE ===');
  Logger.log('Approvers renamed: ' + result.approvers);
  Logger.log('Approvals updated: ' + result.approvals);
  Logger.log('Tokens updated: ' + result.tokens);
  Logger.log('Deviations updated: ' + result.deviations);
  Logger.log('Full result: ' + JSON.stringify(result, null, 2));
  
  return result;
}

function repairSheets() {
  const result = { approvers: [], tokens: [] };
  
  // ── Repair Approvers sheet ──
  const aprSheet = getSheet(SH.APPROVERS, true);
  const aprData  = aprSheet.getDataRange().getValues();
  
  if (aprData.length > 1) {
    // Force entire sheet to text format permanently
    aprSheet.getRange(1, 1, aprSheet.getMaxRows(), aprData[0].length).setNumberFormat('@');
    
    const headers  = aprData[0].map(h => String(h));
    const idCol    = headers.indexOf('id');
    const emailCol = headers.indexOf('email');
    
    if (idCol < 0) {
      Logger.log('Approvers sheet has no id column!');
    } else {
      for (let i = 1; i < aprData.length; i++) {
        const row = aprData[i];
        let idVal = row[idCol];
        const emailVal = row[emailCol] || '';
        let needsFix = false;
        let newId = '';
        
        if (idVal instanceof Date) {
          newId = 'aprv' + i;  // position-based
          needsFix = true;
          Logger.log('Approver row ' + i + ': Date-corrupted, setting id=' + newId);
        } else {
          const idStr = String(idVal || '').trim();
          // Detect date-string corruption
          if (/^\d{4}-\d{2}-\d{2}/.test(idStr) || /^\d+\/\d+\/\d{4}/.test(idStr)) {
            newId = 'aprv' + i;
            needsFix = true;
            Logger.log('Approver row ' + i + ': date-string corrupt (' + idStr + '), setting id=' + newId);
          } else if (idStr !== idStr.toLowerCase()) {
            newId = idStr.toLowerCase();
            needsFix = true;
            Logger.log('Approver row ' + i + ': uppercase id (' + idStr + '), setting id=' + newId);
          }
        }
        
        if (needsFix) {
          // Write as apostrophe-prefixed text to prevent re-corruption
          aprSheet.getRange(i + 1, idCol + 1).setNumberFormat('@').setValue("'" + newId);
          result.approvers.push({ row: i + 1, oldId: String(idVal), newId: newId, email: String(emailVal) });
        }
      }
    }
  }
  
  // ── Repair Tokens sheet ──
  const tokSheet = getSheet(SH.TOKENS, false);
  if (tokSheet) {
    const tokData = tokSheet.getDataRange().getValues();
    if (tokData.length > 1) {
      tokSheet.getRange(1, 1, tokSheet.getMaxRows(), tokData[0].length).setNumberFormat('@');
      
      const tHeaders  = tokData[0].map(h => String(h));
      const tIdCol    = tHeaders.indexOf('approverId');
      const tEmailCol = tHeaders.indexOf('approverEmail');
      
      if (tIdCol >= 0) {
        // Build email→id map from approvers
        const approvers = getApprovers();
        const emailToId = {};
        approvers.forEach(a => {
          if (a.email) emailToId[String(a.email).toLowerCase()] = a.id;
        });
        
        for (let j = 1; j < tokData.length; j++) {
          const row = tokData[j];
          let idVal = row[tIdCol];
          const emailVal = String(row[tEmailCol] || '').toLowerCase();
          let newId = '';
          let needsFix = false;
          
          if (idVal instanceof Date) {
            newId = emailToId[emailVal] || ('aprv_unknown_' + j);
            needsFix = true;
          } else {
            const idStr = String(idVal || '').trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(idStr) || /^\d+\/\d+\/\d{4}/.test(idStr)) {
              newId = emailToId[emailVal] || ('aprv_unknown_' + j);
              needsFix = true;
            } else if (idStr !== idStr.toLowerCase()) {
              newId = idStr.toLowerCase();
              needsFix = true;
            }
          }
          
          if (needsFix) {
            tokSheet.getRange(j + 1, tIdCol + 1).setNumberFormat('@').setValue("'" + newId);
            result.tokens.push({ row: j + 1, oldId: String(idVal), newId: newId, email: emailVal });
          }
        }
      }
    }
  }
  
  Logger.log('Repair complete:');
  Logger.log('Approvers fixed: ' + result.approvers.length);
  Logger.log(JSON.stringify(result.approvers, null, 2));
  Logger.log('Tokens fixed: ' + result.tokens.length);
  Logger.log(JSON.stringify(result.tokens, null, 2));
  
  return result;
}

/**
 * Repair function — fixes duplicate or missing devNums in the Deviations sheet.
 * Use this if you have rows with DEV-1001, DEV-1001, DEV-1001 (all same number)
 * because of the bug where the frontend's placeholder devNum was being used directly.
 *
 * Behavior:
 *   1. Finds rows where devNum is missing or duplicated.
 *   2. Re-assigns consecutive numbers starting from max existing valid devNum + 1.
 *   3. Updates Config!nextDevNum to be max devNum + 1 so future creations are correct.
 */
function repairDevNums() {
  const sheet = getSheet(SH.DEVIATIONS, false);
  if (!sheet) {
    Logger.log('Deviations sheet not found');
    return { error: 'sheet not found' };
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('No deviations to repair');
    return { repaired: 0, max: 1000 };
  }
  const headers = data[0].map(h => String(h));
  const devNumCol = headers.indexOf('devNum');
  const submittedCol = headers.indexOf('submittedAt');
  const idCol = headers.indexOf('id');
  
  if (devNumCol < 0) {
    Logger.log('devNum column not found');
    return { error: 'no devNum column' };
  }
  
  // Pass 1: find max valid devNum across all rows (treat duplicates as needing repair)
  const seenNums = {};
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const cellVal = String(data[i][devNumCol] || '').trim();
    const m = cellVal.match(/(\d+)/);
    const num = m ? parseInt(m[1], 10) : null;
    rows.push({
      rowIdx: i + 1,  // 1-based sheet row
      id: data[i][idCol],
      currentDevNum: cellVal,
      num: num,
      submittedAt: submittedCol >= 0 ? data[i][submittedCol] : null
    });
    if (num !== null) {
      seenNums[num] = (seenNums[num] || 0) + 1;
    }
  }
  
  // Sort rows chronologically by submittedAt (oldest first) for stable renumbering
  rows.sort((a, b) => {
    const ta = a.submittedAt ? (a.submittedAt instanceof Date ? a.submittedAt.getTime() : new Date(a.submittedAt).getTime()) : 0;
    const tb = b.submittedAt ? (b.submittedAt instanceof Date ? b.submittedAt.getTime() : new Date(b.submittedAt).getTime()) : 0;
    return ta - tb;
  });
  
  // Identify which numbers are duplicated or missing
  const duplicates = Object.keys(seenNums).filter(n => seenNums[n] > 1).map(n => parseInt(n, 10));
  Logger.log('Duplicate devNums found: ' + JSON.stringify(duplicates));
  
  // Strategy: keep the FIRST occurrence (oldest by submittedAt) of each devNum,
  // renumber subsequent duplicates and any missing ones.
  // Find the highest valid number to start renumbering from.
  let maxNum = 1000;
  Object.keys(seenNums).forEach(n => {
    const num = parseInt(n, 10);
    if (num > maxNum) maxNum = num;
  });
  
  let nextNum = maxNum + 1;
  const claimedNums = {};  // devNum -> first rowIdx that claims it
  const repairs = [];
  
  rows.forEach(r => {
    if (r.num === null) {
      // Missing devNum — assign new
      const newDev = 'DEV-' + nextNum;
      sheet.getRange(r.rowIdx, devNumCol + 1).setValue(newDev);
      repairs.push({ row: r.rowIdx, id: r.id, old: r.currentDevNum, new: newDev, reason: 'missing' });
      nextNum++;
    } else if (claimedNums[r.num]) {
      // Already claimed by an earlier row — this is a duplicate, renumber
      const newDev = 'DEV-' + nextNum;
      sheet.getRange(r.rowIdx, devNumCol + 1).setValue(newDev);
      repairs.push({ row: r.rowIdx, id: r.id, old: r.currentDevNum, new: newDev, reason: 'duplicate' });
      nextNum++;
    } else {
      // First occurrence — keep
      claimedNums[r.num] = r.rowIdx;
    }
  });
  
  // Update Config!nextDevNum
  const configSheet = getSheet(SH.CONFIG, true);
  const cfgData = configSheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < cfgData.length; i++) {
    if (cfgData[i][0] === 'nextDevNum') {
      configSheet.getRange(i + 1, 2).setValue(nextNum);
      updated = true;
      break;
    }
  }
  if (!updated) {
    configSheet.appendRow(['nextDevNum', nextNum]);
  }
  SpreadsheetApp.flush();
  
  Logger.log('=== REPAIR COMPLETE ===');
  Logger.log('Repaired ' + repairs.length + ' rows');
  Logger.log('Config!nextDevNum set to: ' + nextNum);
  Logger.log('Details: ' + JSON.stringify(repairs, null, 2));
  
  try {
    SpreadsheetApp.getUi().alert(
      'Repair complete:\n\n' +
      'Rows repaired: ' + repairs.length + '\n' +
      'Next devNum will be: DEV-' + nextNum + '\n\n' +
      'Check Apps Script logs for details.'
    );
  } catch(e) {}
  
  return { repaired: repairs.length, nextNum: nextNum, details: repairs };
}

function setWebAppUrl() {
  // ↓↓↓ PASTE YOUR WEB APP URL HERE ↓↓↓
  const NEW_URL = 'https://script.google.com/macros/s/AKfycbx56L2cdsDmQk_1fcVP41kuVIBLV1kwCOBotYUJm8XhgDzvmmpV2iN2fBRN3o9Joj0S/exec';
  // ↑↑↑ ─────────────────────────────── ↑↑↑

  if (!NEW_URL || !NEW_URL.includes('exec')) {
    throw new Error('Paste your Web App URL between the quotes above, then run again.');
  }
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', NEW_URL);
  Logger.log('SUCCESS — URL saved: ' + NEW_URL);
  try {
    SpreadsheetApp.getUi().alert('✅ URL saved:\n\n' + NEW_URL + '\n\nYou only need to do this after each new deployment.');
  } catch(e) {}
}

/**
 * Diagnostic — sends a TEST email to every approver + every creation dist list email.
 * Run this from Apps Script editor. Then check inbox of EACH email and check Executions logs.
 * If an email doesn't arrive, the log will tell you whether it was sent or failed.
 */
function runDiagnosticEmails() {
  const approvers = getApprovers();
  const distLists = getDistLists();
  const remaining = MailApp.getRemainingDailyQuota();
  
  Logger.log('=== DIAGNOSTIC START ===');
  Logger.log('Daily email quota remaining: ' + remaining);
  Logger.log('Approvers loaded: ' + approvers.length);
  approvers.forEach(a => Logger.log('  • ' + a.id + ' | ' + a.name + ' | ' + a.email + ' | role=' + a.role + ' | status=' + a.defaultStatus));
  Logger.log('Creation dist list: ' + JSON.stringify(distLists.creation || []));
  Logger.log('Approval dist list: ' + JSON.stringify(distLists.approval || []));
  Logger.log('---');
  
  let sent = 0, failed = 0;
  approvers.forEach((a, i) => {
    if (!a.email) {
      Logger.log('[' + (i+1) + '/' + approvers.length + '] SKIP - no email for ' + a.id);
      return;
    }
    try {
      MailApp.sendEmail({
        to: a.email,
        subject: '[MWAAF DIAGNOSTIC] Test email — please confirm receipt',
        htmlBody: '<p>This is a test email from the MWAAF Deviation System.</p>' +
                  '<p><strong>Recipient:</strong> ' + a.email + '<br>' +
                  '<strong>Approver ID:</strong> ' + a.id + '<br>' +
                  '<strong>Sent at:</strong> ' + new Date().toString() + '</p>' +
                  '<p>If you receive this, email delivery to your address is working correctly.</p>',
        name: 'MWAAF Diagnostic'
      });
      Logger.log('[' + (i+1) + '/' + approvers.length + '] SENT to ' + a.email);
      sent++;
    } catch(err) {
      Logger.log('[' + (i+1) + '/' + approvers.length + '] FAILED to ' + a.email + ': ' + err.message);
      failed++;
    }
    Utilities.sleep(300);
  });
  
  // Also send to dist list
  (distLists.creation || []).forEach((email, i) => {
    if (!email) return;
    try {
      MailApp.sendEmail({
        to: email,
        subject: '[MWAAF DIAGNOSTIC] Dist list test — please confirm receipt',
        htmlBody: '<p>This is a test FYI email from MWAAF Deviation System.</p>' +
                  '<p><strong>Recipient (creation dist list):</strong> ' + email + '<br>' +
                  '<strong>Sent at:</strong> ' + new Date().toString() + '</p>',
        name: 'MWAAF Diagnostic'
      });
      Logger.log('DIST [' + (i+1) + '] SENT to ' + email);
      sent++;
    } catch(err) {
      Logger.log('DIST [' + (i+1) + '] FAILED to ' + email + ': ' + err.message);
      failed++;
    }
    Utilities.sleep(300);
  });
  
  Logger.log('---');
  Logger.log('=== DIAGNOSTIC COMPLETE: sent=' + sent + ', failed=' + failed + ' ===');
  Logger.log('Quota remaining now: ' + MailApp.getRemainingDailyQuota());
  
  try {
    SpreadsheetApp.getUi().alert('Diagnostic complete:\n\n' +
      'Sent: ' + sent + '\nFailed: ' + failed + '\n' +
      'Quota remaining: ' + MailApp.getRemainingDailyQuota() + '/100\n\n' +
      'Check each email inbox (including spam) and review Executions logs for details.');
  } catch(e) {}
  
  return { sent: sent, failed: failed, quota: MailApp.getRemainingDailyQuota() };
}

/** Test function — sends a test email to yourself. */
function testEmail() {
  const email = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to:       email,
    subject:  'MWAAF Test Email',
    htmlBody: emailWrapper(`<div class="hdr"><div class="hdr-title">✅ Test Email</div><div class="hdr-sub">If you see this, email is working!</div></div><div class="body"><p>Apps Script email is configured correctly.</p></div>`),
    name:     'MWAAF Deviation System',
  });
  Logger.log('Test email sent successfully to: ' + email);
  try { SpreadsheetApp.getUi().alert('Test email sent to ' + email); } catch(e) {}
}

/** Adds a custom menu to the spreadsheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Deviation System')
    .addItem('1. Setup All Sheets', 'setupSheets')
    .addItem('2. Set Web App URL', 'setWebAppUrl')
    .addItem('3. Test Email (to yourself)', 'testEmail')
    .addItem('4. Diagnostic — Email ALL approvers', 'runDiagnosticEmails')
    .addSeparator()
    .addItem('🔧 Repair: Fix duplicate DEV-numbers', 'repairDevNums')
    .addSeparator()
    .addItem('View Script Logs', 'openLogs')
    .addToUi();
}

function openLogs() {
  SpreadsheetApp.getUi().alert('Open: Extensions → Apps Script → Executions to see logs.');
}
