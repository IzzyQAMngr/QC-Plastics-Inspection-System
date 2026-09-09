/*************************************************************
 * PLASTICS DROP FREEZE TESTING — web app backend
 * Rebuilt 2026-08-07 onto the Run picker (Runs replace the old ad-hoc Batch concept —
 * see Shared.gs) and the Spec Register's Functional Tests tab (auto-fills the test
 * method/acceptance criteria/sample size once a Drop Freeze test is picked for the
 * Run's mold; scoped to Test Name containing "Drop Freeze" — the register's other
 * fit tests, e.g. Gauge Fit/Cover Fit/Handle Fit, aren't in scope for this module).
 *************************************************************/

function getDropFreezeLogSheet_() {
  const sheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + DROPFREEZE_LOG_SHEET_NAME + '" sheet not found. Run oneTimeSetup() first.');
  return sheet;
}

/**
 * Called by DropFreezeView.html/OpenSamplesView.html on load to populate dropdowns — split into
 * two smaller calls (see fetchDropFreezeFormData_ in Combobox.html) rather than one combined
 * {runs, angleOptions, inspectors, shifts, openRecords} response. The combined shape reliably
 * lost its google.script.run callback over the sandboxed iframe bridge — confirmed 2026-09-09 by
 * direct testing: each half below completes normally every time on its own, but returning both
 * together from one call never did, regardless of timing, caching, or the function's name.
 */
function getDropFreezeBaseFormData() {
  return {
    runs: getActiveRuns_(),
    angleOptions: DROP_ANGLE_OPTIONS,
    inspectors: getInspectorList_(),
    shifts: getShiftList_(),
  };
}

function getOpenDropFreezeRecords() { return listOpenDropFreezeRecords_(); }

/** Drop Freeze test protocols defined for a mold — drives the read-only reference box
 *  (method/acceptance criteria/sample size/equipment) once a test is picked. */
function getDropFreezeTestsForMold(moldId) {
  return getFunctionalTestsForMold_(moldId, 'Drop Freeze');
}

/** Cavity picker options for a mold — same register lookup In-Process uses to auto-generate its cavity rows. */
function getCavityIdsForMold(moldId) {
  return getCavityIds_(moldId);
}

/** The Spec Register's Functional Tests tab has no separate duration/conditioning-time column —
 *  the hold time only ever appears written into the test's own name, e.g. "Drop Freeze Test —
 *  24 hr" or "— 72 hr" (see getFunctionalTestsForMold_) — so that text is the only place this
 *  lives today. Returns null if the name has no recognizable "<number> hr" pattern, rather than
 *  guessing a default — callers must treat that as "no due date computable", not "on time". */
function parseDropFreezeDurationHours_(testName) {
  const m = String(testName || '').match(/(\d+(?:\.\d+)?)\s*hr/i);
  return m ? parseFloat(m[1]) : null;
}

/**
 * Every Drop Freeze record still holding at least one OPEN sample, with enough context
 * (Run/Mold/Cavity/Test/dates/how many samples, plus a computed due date) to identify and
 * prioritize it without opening the sheet — drives both the Test Results tab's dropdown and the
 * Open Samples dashboard.
 *
 * Reads the log in two passes instead of pulling every historical row's full width just to find
 * the usually-much-smaller set of still-open ones (this was the main reason Open Samples got
 * slow to load as the log grew): a narrow RecordKey+Status-only scan across the whole sheet
 * first locates, for every record with at least one open sample, where that record's block of
 * rows STARTS (not where its first open row happens to be — a record with some already-completed
 * line items ahead of its still-open one would otherwise be undercounted) — then a single
 * full-width read covers from the earliest such start to the end of the sheet. Skips the
 * full-width read entirely once there are no open records at all. Deliberately does NOT bound
 * this by date/age the way Dashboard.gs's In-Process read does — an unusually old still-open
 * record is exactly what the Past Due KPI exists to surface, not hide.
 */
function listOpenDropFreezeRecords_() {
  const sheet = getDropFreezeLogSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const keyCol = headers.indexOf('RecordKey'), statusCol = headers.indexOf('Status');
  if (keyCol < 0 || statusCol < 0) return [];

  const n = lastRow - 1;
  const keys = sheet.getRange(2, keyCol + 1, n, 1).getValues().map(r => String(r[0] || '').trim());
  const statuses = sheet.getRange(2, statusCol + 1, n, 1).getValues().map(r => String(r[0] || '').trim().toUpperCase());

  const blockStart = {}, hasOpen = {};
  for (let i = 0; i < n; i++) {
    const k = keys[i];
    if (!k) continue;
    if (!(k in blockStart)) blockStart[k] = i;
    if (statuses[i] === 'OPEN') hasOpen[k] = true;
  }
  const openKeys = Object.keys(hasOpen);
  if (openKeys.length === 0) return [];

  const startIdx = Math.min.apply(null, openKeys.map(k => blockStart[k]));
  const startRow = 2 + startIdx;
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  const rows = data.map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });

  const groups = new Map();
  rows.forEach(r => {
    const key = String(r.RecordKey || '').trim();
    if (!key || !hasOpen[key]) return;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  const records = [];
  groups.forEach((groupRows, key) => {
    const openCount = groupRows.filter(r => String(r.Status || '').trim().toUpperCase() === 'OPEN').length;
    if (openCount === 0) return;
    const first = groupRows[0];
    const createdAt = first.Created instanceof Date ? first.Created : new Date(first.Created);
    const durationHours = parseDropFreezeDurationHours_(first['Test Name']);
    const dueAt = (durationHours !== null && !isNaN(createdAt.getTime()))
      ? new Date(createdAt.getTime() + durationHours * 3600000) : null;
    records.push({
      recordKey: key,
      runId: first['Run ID'] || '', line: first['Line #'] || '',
      moldId: first['Mold ID'] || '', moldDescription: first['Mold Description'] || '',
      itemNo: first['Item No'] || '', customerName: first['Customer Name'] || '',
      cavity: first.Cavity || '', testName: first['Test Name'] || '',
      dateOfMfg: dateToStr_(first.DateOfMfg),
      createdAt: dateToStr_(createdAt), durationHours: durationHours, dueAt: dateToStr_(dueAt),
      totalSamples: groupRows.length, openSamples: openCount,
    });
  });
  records.sort((a, b) => b.recordKey.localeCompare(a.recordKey));
  return records;
}

/** Loads a saved packet's line items for editing. */
function loadDropFreezeRecord(recordKey) {
  const sheet = getDropFreezeLogSheet_();
  const rows = readSheetObjects_(sheet)
    .filter(r => String(r.RecordKey || '').trim() === String(recordKey).trim())
    .sort((a, b) => Number(a.LineItem) - Number(b.LineItem));
  if (rows.length === 0) throw new Error('No line items found for record: ' + recordKey);
  return {
    recordKey: recordKey,
    lineItems: rows.map(r => ({
      runId: r['Run ID'] || '', line: r['Line #'] || '', shift: r.Shift || '', customerName: r['Customer Name'] || '',
      moldId: r['Mold ID'] || '', moldDescription: r['Mold Description'] || '', productType: r['Product Type'] || '',
      resinLot: r['Resin Lot'] || '', itemNo: r['Item No'] || '', itemDescription: r['Item Description'] || '',
      cavity: r.Cavity || '', testName: r['Test Name'] || '',
      dateOfMfg: dateToStr_(r.DateOfMfg), testDate: dateToStr_(r.TestDate), testedBy: r.TestedBy || '',
      sampleNo: r.SampleNo || '', sampleCount: r.SampleCount || '',
      freezerTemp: r.FreezerTemp, dropHeight: r.DropHeight, dropAngle: r.DropAngle || '', result: r.Result || '',
      failureDescription: r.FailureDescription || '', notes: r.Notes || '',
    })),
  };
}

function makeDailyRecordKey_(sheet, dateOfMfgDisplay) {
  const parsed = new Date(dateOfMfgDisplay);
  const baseDate = isNaN(parsed.getTime()) ? new Date() : parsed;
  return makeSequentialId_(sheet, 'RecordKey', 'QC', baseDate);
}

/** Builds one Drop Freeze log row. Run context (Line #, Mold, Product Type, Resin Lot, Item,
 *  Customer Name) always comes from the resolved Run — never trusted from the client — same as
 *  In-Process/Start-Up. Status is derived from whether `li.result` is filled in, so a blank
 *  result (samples just loaded, not tested yet) always lands as OPEN. */
function buildDropFreezeRow_(run, li, recordKey, lineItem, now, tz) {
  const resultRaw = String(li.result || '').trim();
  const normalized = resultRaw.replace(/[^\w\s]/g, '').toUpperCase();
  const status = (normalized.includes('PASS') || normalized.includes('FAIL') || normalized.includes('INCONCLUSIVE'))
    ? 'COMPLETE' : 'OPEN';
  const d = li.dateOfMfg ? new Date(li.dateOfMfg) : now;
  return {
    RecordKey: recordKey, LineItem: lineItem, Status: status, Created: now, Updated: now,
    'Run ID': run.runId, 'Line #': run.line, Shift: li.shift || run.shift, 'Customer Name': run.customerName,
    'Mold ID': run.moldId, 'Mold Description': run.moldDescription, 'Product Type': run.productType,
    'Resin Lot': run.resinLot, 'Item No': run.item, 'Item Description': run.itemDescription,
    Cavity: li.cavity || '', 'Test Name': li.testName || '',
    DateOfMfg: li.dateOfMfg || '', TestDate: li.testDate || '', TestedBy: li.testedBy || '',
    SampleNo: li.sampleNo || '', SampleCount: li.sampleCount || '',
    FreezerTemp: li.freezerTemp, DropHeight: li.dropHeight, DropAngle: li.dropAngle || '', Result: li.result || '',
    FailureDescription: li.failureDescription || '', Notes: li.notes || '',
    Month: Utilities.formatDate(d, tz, 'MMMM'), Year: Utilities.formatDate(d, tz, 'yyyy'),
  };
}

/**
 * Saves (creates or replaces) a Drop Freeze packet.
 * payload: { recordKey (nullable), lineItems: [{runId, cavity, testName, dateOfMfg, testDate,
 *   testedBy, shift, freezerTemp, dropHeight, dropAngle, result, failureDescription, notes}] }
 */
function saveDropFreezePacket(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getDropFreezeLogSheet_();
    ensureColumnExists_(sheet, 'SampleNo');
    ensureColumnExists_(sheet, 'SampleCount');
    const items = payload.lineItems || [];
    const active = items.filter(li => String(li.runId || '').trim());
    if (active.length === 0) throw new Error('No active line items — a Run is required on at least one line.');

    let recordKey = String(payload.recordKey || '').trim();
    if (recordKey && !/^QC-\d{6}-\d+$/.test(recordKey)) {
      throw new Error('Invalid QC Record #: "' + recordKey + '"');
    }
    if (!recordKey) recordKey = makeDailyRecordKey_(sheet, active[0].dateOfMfg);

    // Replace any existing rows for this record (edit-in-place)
    deleteRowsWhere_(sheet, 'RecordKey', recordKey);

    const tz = getDb_().getSpreadsheetTimeZone();
    const now = new Date();
    const rowsToAppend = active.map((li, idx) => {
      const run = getRun_(li.runId);
      if (!run) throw new Error('Run not found: ' + li.runId);
      return buildDropFreezeRow_(run, li, recordKey, idx + 1, now, tz);
    });

    appendObjectsAsRows_(sheet, rowsToAppend);
    try { sendDropFreezeFailEmails_(recordKey, rowsToAppend); } catch (e) { /* best-effort */ }

    return { recordKey: recordKey, savedCount: rowsToAppend.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Logs samples for EVERY cavity of the mold at once — one new Open record per cavity, each
 * holding sampleCount blank-result line items — instead of making the tech pick a single cavity
 * and repeat "Log Samples" per cavity. Mirrors In-Process auto-generating a row per cavity.
 * payload: { runId, shift, testName, dateOfMfg, cavities: [cavityId, ...], sampleCount }
 */
function logDropFreezeSamples(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getDropFreezeLogSheet_();
    ensureColumnExists_(sheet, 'SampleNo');
    ensureColumnExists_(sheet, 'SampleCount');

    const run = getRun_(payload.runId);
    if (!run) throw new Error('Run not found: ' + payload.runId);

    const cavities = (payload.cavities || []).map(c => String(c || '').trim()).filter(Boolean);
    if (cavities.length === 0) throw new Error('No cavities to log — this mold has none defined in the register.');
    const count = Math.max(1, Number(payload.sampleCount) || 1);

    const tz = getDb_().getSpreadsheetTimeZone();
    const now = new Date();
    const results = [];

    cavities.forEach(cavity => {
      // Computed fresh per cavity so each just-appended record's rows count toward the next
      // cavity's sequence number — see makeSequentialId_.
      const recordKey = makeDailyRecordKey_(sheet, payload.dateOfMfg);
      const rows = [];
      for (let i = 0; i < count; i++) {
        const li = {
          shift: payload.shift, testName: payload.testName, dateOfMfg: payload.dateOfMfg, cavity: cavity,
          sampleNo: count > 1 ? (i + 1) : '', sampleCount: count > 1 ? count : '',
        };
        rows.push(buildDropFreezeRow_(run, li, recordKey, i + 1, now, tz));
      }
      appendObjectsAsRows_(sheet, rows);
      results.push({ cavity: cavity, recordKey: recordKey, savedCount: rows.length });
    });

    return {
      results: results,
      totalRecords: results.length,
      totalSamples: results.reduce((sum, r) => sum + r.savedCount, 0),
    };
  } finally {
    lock.releaseLock();
  }
}

function sendDropFreezeFailEmails_(recordKey, rows) {
  const emails = getNotificationEmails_();
  if (emails.length === 0) return;
  const props = PropertiesService.getScriptProperties();
  const dbId = getDb_().getId();

  rows.forEach((row, i) => {
    const lineItem = i + 1;
    const dedupeKey = 'DF_FAIL_SAVED__' + dbId + '__' + recordKey + '__LINE_' + lineItem;
    if (!isFailValue_(row.Result)) { props.deleteProperty(dedupeKey); return; }
    if (props.getProperty(dedupeKey)) return;

    const labeled = Object.keys(row)
      .filter(k => ['RecordKey', 'LineItem', 'Status', 'Created', 'Updated'].indexOf(k) === -1)
      .filter(k => String(row[k] || '').trim() !== '')
      .map(k => k + ': ' + row[k]);

    const subject = 'QC Drop Freeze Test FAIL Saved — ' + recordKey + ' (Line ' + lineItem + ')';
    const body = 'A FAIL result was saved.\n\nRecord #: ' + recordKey + '\nLine Item #: ' + lineItem +
      '\n\n--------------------------------\nLINE ITEM DETAILS\n--------------------------------\n' +
      labeled.join('\n') + '\n\nOpen the CSC QC Inspection System app to view/edit this record.';

    emails.forEach(to => MailApp.sendEmail({ to: to, subject: subject, body: body }));
    props.setProperty(dedupeKey, new Date().toISOString());
  });
}
