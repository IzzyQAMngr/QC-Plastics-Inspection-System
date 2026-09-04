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

/** Called by DropFreezeView.html on load to populate dropdowns. */
function getDropFreezeFormData() {
  return {
    runs: getActiveRuns_(),
    angleOptions: DROP_ANGLE_OPTIONS,
    inspectors: getInspectorList_(),
    shifts: getShiftList_(),
    openRecords: listOpenDropFreezeRecords_(),
  };
}

/** Drop Freeze test protocols defined for a mold — drives the read-only reference box
 *  (method/acceptance criteria/sample size/equipment) once a test is picked. */
function getDropFreezeTestsForMold(moldId) {
  return getFunctionalTestsForMold_(moldId, 'Drop Freeze');
}

/** Cavity picker options for a mold — same register lookup In-Process uses to auto-generate its cavity rows. */
function getCavityIdsForMold(moldId) {
  return getCavityIds_(moldId);
}

function listOpenDropFreezeRecords_() {
  const sheet = getDropFreezeLogSheet_();
  const rows = readSheetObjects_(sheet);
  const seen = new Set();
  const records = [];
  rows.forEach(r => {
    const key = String(r.RecordKey || '').trim();
    const status = String(r.Status || '').trim().toUpperCase();
    if (key && status === 'OPEN' && !seen.has(key)) { seen.add(key); records.push(key); }
  });
  records.sort((a, b) => b.localeCompare(a));
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

/**
 * Saves (creates or replaces) a Drop Freeze packet.
 * payload: { recordKey (nullable), lineItems: [{runId, cavity, testName, dateOfMfg, testDate,
 *   testedBy, shift, freezerTemp, dropHeight, dropAngle, result, failureDescription, notes}] }
 * Each line's Run context (Line #, Mold, Product Type, Resin Lot, Item, Customer Name) is
 * resolved server-side from the Run — never trusted from the client — same as In-Process/Start-Up.
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
      const resultRaw = String(li.result || '').trim();
      const normalized = resultRaw.replace(/[^\w\s]/g, '').toUpperCase();
      const status = (normalized.includes('PASS') || normalized.includes('FAIL') || normalized.includes('INCONCLUSIVE'))
        ? 'COMPLETE' : 'OPEN';
      const d = li.dateOfMfg ? new Date(li.dateOfMfg) : now;
      return {
        RecordKey: recordKey, LineItem: idx + 1, Status: status, Created: now, Updated: now,
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
    });

    appendObjectsAsRows_(sheet, rowsToAppend);
    try { sendDropFreezeFailEmails_(recordKey, rowsToAppend); } catch (e) { /* best-effort */ }

    return { recordKey: recordKey, savedCount: rowsToAppend.length };
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
