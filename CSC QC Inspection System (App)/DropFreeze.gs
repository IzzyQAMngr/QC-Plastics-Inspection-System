/*************************************************************
 * PLASTICS DROP FREEZE TESTING — web app backend
 * Ported from legacy Plastics Drop Freeze Testing/Code.js
 * (submitQCCore_ / loadQcPacketFromDropdownCore_ / sendFailEmailsForSavedPacket_)
 *************************************************************/

function getDropFreezeLogSheet_() {
  const sheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + DROPFREEZE_LOG_SHEET_NAME + '" sheet not found. Run oneTimeSetup() first.');
  return sheet;
}

/** Called by DropFreezeView.html on load to populate dropdowns. */
function getDropFreezeFormData() {
  return {
    molds: getAllMoldsList_(),         // [{moldId, description, productType}] — Tool Code combobox; picking one auto-fills description+productType client-side
    itemList: getItemList_(),          // [{itemNo, description}] — Item No. combobox
    angleOptions: DROP_ANGLE_OPTIONS,
    inspectors: getInspectorList_(),
    shifts: getShiftList_(),
    openBatches: getOpenBatches_(),
    openRecords: listOpenDropFreezeRecords_(),
  };
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
    batchId: rows[0].BatchID || '',
    lineItems: rows.map(r => ({
      lineNum: r.LineNum, dateOfMfg: r.DateOfMfg, shift: r.Shift, pallet: r.Pallet,
      toolCode: r.ToolCode, productType: r.ProductType, cavity: r.Cavity, resinId: r.ResinID, testType: r.TestType,
      itemNo: r.ItemNo, itemDescription: r.ItemDescription,
      testDate: r.TestDate, testedBy: r.TestedBy, freezerTemp: r.FreezerTemp,
      dropHeight: r.DropHeight, dropAngle: r.DropAngle, result: r.Result,
      failureDescription: r.FailureDescription, notes: r.Notes, batchId: r.BatchID,
    })),
  };
}

function makeDailyRecordKey_(sheet, dateOfMfgDisplay) {
  const tz = getDb_().getSpreadsheetTimeZone();
  let baseDate = new Date();
  const parsed = new Date(dateOfMfgDisplay);
  if (!isNaN(parsed.getTime())) baseDate = parsed;
  const dateStr = Utilities.formatDate(baseDate, tz, 'yyyyMMdd');
  const rows = readSheetObjects_(sheet);
  let maxSeq = 0;
  rows.forEach(r => {
    const m = String(r.RecordKey || '').match(/^QC-(\d{8})-(\d{3})$/);
    if (m && m[1] === dateStr) {
      const seq = Number(m[2]);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return 'QC-' + dateStr + '-' + String(maxSeq + 1).padStart(3, '0');
}

/**
 * Saves (creates or replaces) a Drop Freeze packet.
 * payload: { recordKey (nullable), lineItems: [{lineNum,dateOfMfg,shift,pallet,toolCode,cavity,
 *   resinId,testType,testDate,testedBy,freezerTemp,dropHeight,dropAngle,result,failureDescription,
 *   notes,batchId}] }
 */
function saveDropFreezePacket(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getDropFreezeLogSheet_();
    const items = payload.lineItems || [];
    const active = items.filter(li => String(li.toolCode || '').trim());
    if (active.length === 0) throw new Error('No active line items — Tool Code is required on at least one line.');

    let recordKey = String(payload.recordKey || '').trim();
    if (recordKey && !/^QC-\d{8}-\d{3}$/.test(recordKey)) {
      throw new Error('Invalid QC Record #: "' + recordKey + '"');
    }
    if (!recordKey) recordKey = makeDailyRecordKey_(sheet, active[0].dateOfMfg);

    // Replace any existing rows for this record (edit-in-place)
    deleteRowsWhere_(sheet, 'RecordKey', recordKey);

    const now = new Date();
    const rowsToAppend = active.map((li, idx) => {
      const resultRaw = String(li.result || '').trim();
      const normalized = resultRaw.replace(/[^\w\s]/g, '').toUpperCase();
      const status = (normalized.includes('PASS') || normalized.includes('FAIL') || normalized.includes('INCONCLUSIVE'))
        ? 'COMPLETE' : 'OPEN';
      return {
        RecordKey: recordKey, LineItem: idx + 1, Status: status, Created: now, Updated: now,
        LineNum: li.lineNum, DateOfMfg: li.dateOfMfg, Shift: li.shift, Pallet: li.pallet,
        ToolCode: li.toolCode, ProductType: li.productType, Cavity: li.cavity, ResinID: li.resinId, TestType: li.testType,
        ItemNo: li.itemNo, ItemDescription: li.itemDescription,
        TestDate: li.testDate, TestedBy: li.testedBy, FreezerTemp: li.freezerTemp,
        DropHeight: li.dropHeight, DropAngle: li.dropAngle, Result: li.result,
        FailureDescription: li.failureDescription, Notes: li.notes, BatchID: li.batchId || payload.batchId || '',
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
