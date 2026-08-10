/*************************************************************
 * START-UP VERIFICATION — Run-linked checklist that qualifies a Run.
 * The Log is a flat/long sheet: one row per checklist item per submission, all sharing the
 * same Verification Record # and header context (mirrors the existing QC Database pattern).
 * Deviation sign-off (Deviation Approved by / PFA Signed off by / Supervisor Authorization /
 * PFA ID) is never written by the submitter directly — it's appended later by the Deviation
 * Approvals flow below, and a PFA cannot be signed off until the deviation is approved first.
 *************************************************************/

function getSuLogSheet_() {
  const sheet = getDb_().getSheetByName(SU_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + SU_LOG_SHEET_NAME + '" sheet not found.');
  return sheet;
}

function makeVerificationRecordId_() {
  const sheet = getSuLogSheet_();
  const tz = getDb_().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf('Verification Record #');
    if (col >= 0) {
      const ids = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues().flat();
      ids.forEach(id => {
        const m = String(id || '').match(/^SUV-(\d{8})-(\d{3})$/);
        if (m && m[1] === dateStr) { const seq = Number(m[2]); if (seq > maxSeq) maxSeq = seq; }
      });
    }
  }
  return 'SUV-' + dateStr + '-' + String(maxSeq + 1).padStart(3, '0');
}

function makePfaId_() {
  const sheet = getSuLogSheet_();
  const tz = getDb_().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf('Actual Value');
    if (col >= 0) {
      const vals = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues().flat();
      vals.forEach(v => {
        const m = String(v || '').match(/^PFA-(\d{8})-(\d{3})$/);
        if (m && m[1] === dateStr) { const seq = Number(m[2]); if (seq > maxSeq) maxSeq = seq; }
      });
    }
  }
  return 'PFA-' + dateStr + '-' + String(maxSeq + 1).padStart(3, '0');
}

function getRecordRows_(recordId) {
  return readSheetObjects_(getSuLogSheet_()).filter(r => String(r['Verification Record #'] || '').trim() === String(recordId).trim());
}

/** Shared header columns to carry onto every new row appended for a record (initial or later sign-off rows). */
function contextFromRow_(row) {
  return {
    'Verification Record #': row['Verification Record #'] || '', 'Timestamp saved': dateToStr_(new Date()),
    'Run ID': row['Run ID'] || '', 'Verification Date': row['Verification Date'] || '', 'Verification Time': row['Verification Time'] || '',
    'QC Tech Name': row['QC Tech Name'] || '', 'Shift': row['Shift'] || '', 'Foreman': row['Foreman'] || '',
    'Start-Up Tech': row['Start-Up Tech'] || '', 'Line #': row['Line #'] || '', 'Run Qty': row['Run Qty'] || '',
    'Customer Name': row['Customer Name'] || '', 'Mold ID': row['Mold ID'] || '', 'Mold Description': row['Mold Description'] || '',
    'Item': row['Item'] || '', 'Item Description': row['Item Description'] || '', 'Month': row['Month'] || '', 'Year': row['Year'] || '',
  };
}

function computeItemStatus_(valueType, value) {
  if (String(valueType || '').replace(/\s+/g, '').toLowerCase() === 'yes/no') {
    const v = String(value || '').trim();
    if (v === 'Yes') return 'Pass';
    if (v === 'No') return 'Fail';
  }
  return '';
}

/** Called by StartUpVerificationView.html on load. */
function getStartUpVerificationFormData() {
  return {
    items: getStartUpItemsList_(),
    runs: getActiveRuns_(),
    qcTechOptions: getInspectorList_(),
    startUpTechOptions: getStartUpTechList_(),
    shiftOptions: getShiftList_(),
    foremanOptions: getForemanList_(),
    pfaAuthOptions: getDeviationAuthList_(),
  };
}

/**
 * payload: {runId, verificationDate, verificationTime, qcTechName, startUpTechName, shift, foreman,
 *   checklist: [{item, valueType, unit, value, notes}], pfaDirectSignOff}
 * pfaDirectSignOff is only honored when the checklist's "Was there a deviation?" answer is not "Yes" —
 * that's the self-serve close-out path (nothing to approve, so no lock, no email).
 */
function saveStartUpVerification_(payload) {
  const run = getRun_(payload.runId);
  if (!run) throw new Error('Run not found: ' + payload.runId);

  const tz = getDb_().getSpreadsheetTimeZone();
  const d = payload.verificationDate ? new Date(payload.verificationDate) : new Date();
  const recordId = makeVerificationRecordId_();

  const context = {
    'Verification Record #': recordId, 'Timestamp saved': dateToStr_(new Date()), 'Run ID': run.runId,
    'Verification Date': payload.verificationDate || '', 'Verification Time': payload.verificationTime || '',
    'QC Tech Name': payload.qcTechName || '', 'Shift': payload.shift || '', 'Foreman': payload.foreman || '',
    'Start-Up Tech': payload.startUpTechName || '', 'Line #': run.line || '', 'Run Qty': run.runQty || '',
    'Customer Name': run.customerName || '', 'Mold ID': run.moldId || '', 'Mold Description': run.moldDescription || '',
    'Item': run.item || '', 'Item Description': run.itemDescription || '',
    'Month': Utilities.formatDate(d, tz, 'MMMM'), 'Year': Utilities.formatDate(d, tz, 'yyyy'),
  };

  const checklist = payload.checklist || [];
  const hasDeviation = checklist.some(c => c.item === 'Was there a deviation?' && String(c.value).trim() === 'Yes');

  const rows = checklist.map(c => Object.assign({}, context, {
    'Verification Item': c.item, 'Value Type': c.valueType || '', 'Unit': c.unit || '',
    'Actual Value': (c.value === undefined || c.value === null) ? '' : c.value,
    'Status': computeItemStatus_(c.valueType, c.value), 'Notes': c.notes || '',
  }));

  let qualified = false;
  if (!hasDeviation && payload.pfaDirectSignOff) {
    const pfaId = makePfaId_();
    rows.push(Object.assign({}, context, { 'Verification Item': 'PFA Signed off by', 'Value Type': 'Drop Down', 'Unit': '', 'Actual Value': payload.pfaDirectSignOff, 'Status': '', 'Notes': '' }));
    rows.push(Object.assign({}, context, { 'Verification Item': 'PFA completed (PFA ID)', 'Value Type': 'Autofill', 'Unit': '', 'Actual Value': pfaId, 'Status': '', 'Notes': '' }));
    qualifyRun_(run.runId);
    qualified = true;
  }

  appendObjectsAsRows_(getSuLogSheet_(), rows);
  if (hasDeviation) sendDeviationEmail_(recordId, run);

  return { recordId: recordId, qualified: qualified, hasDeviation: hasDeviation };
}
function saveStartUpVerification(payload) { return saveStartUpVerification_(payload); }

function sendDeviationEmail_(recordId, run) {
  const emails = getNotificationEmails_();
  if (!emails.length) return;
  const link = WEB_APP_URL + '?view=devapproval&record=' + encodeURIComponent(recordId);
  const subject = 'Deviation needs approval — Run ' + run.runId;
  const body = 'A deviation was reported during Start-Up Verification.\n\n' +
    'Run: ' + run.runId + '\n' +
    'Mold: ' + run.moldId + (run.moldDescription ? ' — ' + run.moldDescription : '') + '\n' +
    'Line: ' + run.line + '\n' +
    'Record #: ' + recordId + '\n\n' +
    'Review and approve here:\n' + link + '\n';
  MailApp.sendEmail({ to: emails.join(','), subject: subject, body: body });
}

// ================= DEVIATION APPROVALS =================

/** Deviations still needing action — approval, then a PFA sign-off. Drops off once both are done. */
function getPendingDeviations_() {
  const rows = readSheetObjects_(getSuLogSheet_());
  const byRecord = {};
  rows.forEach(r => {
    const id = r['Verification Record #'];
    if (!id) return;
    (byRecord[id] = byRecord[id] || []).push(r);
  });
  const out = [];
  Object.keys(byRecord).forEach(id => {
    const group = byRecord[id];
    const hasDeviation = group.some(r => r['Verification Item'] === 'Was there a deviation?' && String(r['Actual Value']).trim() === 'Yes');
    if (!hasDeviation) return;
    const approved = group.some(r => r['Verification Item'] === 'Deviation Approved by' && String(r['Actual Value']).trim());
    const signedOff = group.some(r => r['Verification Item'] === 'PFA Signed off by' && String(r['Actual Value']).trim());
    if (approved && signedOff) return;
    const ctx = group[0];
    const descRow = group.find(r => r['Verification Item'] === 'Deviation Description / Notes');
    out.push({
      recordId: id, runId: ctx['Run ID'] || '', moldId: ctx['Mold ID'] || '', moldDescription: ctx['Mold Description'] || '',
      line: ctx['Line #'] || '', verificationDate: dateToStr_(ctx['Verification Date']), qcTechName: ctx['QC Tech Name'] || '',
      startUpTechName: ctx['Start-Up Tech'] || '', deviationDescription: descRow ? String(descRow['Actual Value'] || '') : '',
      status: approved ? 'Awaiting PFA Sign-off' : 'Awaiting Approval',
    });
  });
  return out;
}
function getPendingDeviations() { return getPendingDeviations_(); }

/** Called by DeviationApprovalView.html on load. */
function getDeviationApprovalFormData() {
  return { pending: getPendingDeviations_(), authOptions: getDeviationAuthList_() };
}

function getDeviationRecordDetail_(recordId) {
  const rows = getRecordRows_(recordId);
  if (!rows.length) return null;
  const ctx = rows[0];
  const findVal = item => { const r = rows.find(x => x['Verification Item'] === item); return r ? String(r['Actual Value'] || '') : ''; };
  return {
    recordId: recordId, runId: ctx['Run ID'] || '', moldId: ctx['Mold ID'] || '', moldDescription: ctx['Mold Description'] || '',
    line: ctx['Line #'] || '', item: ctx['Item'] || '', itemDescription: ctx['Item Description'] || '',
    verificationDate: dateToStr_(ctx['Verification Date']), verificationTime: ctx['Verification Time'] || '',
    qcTechName: ctx['QC Tech Name'] || '', startUpTechName: ctx['Start-Up Tech'] || '',
    deviationDescription: findVal('Deviation Description / Notes'),
    deviationApprovedBy: findVal('Deviation Approved by'),
    supervisorAuthorization: findVal('Supervisor Authorization (if deviation)'),
    pfaSignedOffBy: findVal('PFA Signed off by'),
    pfaId: findVal('PFA completed (PFA ID)'),
  };
}
function getDeviationRecordDetail(recordId) { return getDeviationRecordDetail_(recordId); }

/** Approves the deviation — does NOT touch PFA Signed off by, which stays locked until this exists. */
function approveDeviation_(recordId, approverName, supervisorName) {
  const rows = getRecordRows_(recordId);
  if (!rows.length) throw new Error('Verification record not found: ' + recordId);
  const authList = getDeviationAuthList_();
  if (authList.indexOf(approverName) < 0) throw new Error(approverName + ' is not on the Deviation Authorization List.');
  if (supervisorName && authList.indexOf(supervisorName) < 0) throw new Error(supervisorName + ' is not on the Deviation Authorization List.');

  const context = contextFromRow_(rows[0]);
  const newRows = [Object.assign({}, context, { 'Verification Item': 'Deviation Approved by', 'Value Type': 'Drop Down', 'Unit': '', 'Actual Value': approverName, 'Status': '', 'Notes': '' })];
  if (supervisorName) newRows.push(Object.assign({}, context, { 'Verification Item': 'Supervisor Authorization (if deviation)', 'Value Type': 'Drop Down', 'Unit': '', 'Actual Value': supervisorName, 'Status': '', 'Notes': '' }));
  appendObjectsAsRows_(getSuLogSheet_(), newRows);
  return getDeviationRecordDetail_(recordId);
}
function approveDeviation(recordId, approverName, supervisorName) { return approveDeviation_(recordId, approverName, supervisorName); }

/** Signs off the PFA and qualifies the Run — blocked until the deviation has been approved. */
function signPfa_(recordId, qcName, adjustmentsText) {
  const detail = getDeviationRecordDetail_(recordId);
  if (!detail) throw new Error('Verification record not found: ' + recordId);
  if (!detail.deviationApprovedBy) throw new Error('This deviation has not been approved yet — the PFA cannot be signed off until it is.');

  const rows = getRecordRows_(recordId);
  const context = contextFromRow_(rows[0]);
  const pfaId = makePfaId_();
  const newRows = [
    Object.assign({}, context, { 'Verification Item': 'PFA Signed off by', 'Value Type': 'Drop Down', 'Unit': '', 'Actual Value': qcName, 'Status': '', 'Notes': '' }),
    Object.assign({}, context, { 'Verification Item': 'PFA completed (PFA ID)', 'Value Type': 'Autofill', 'Unit': '', 'Actual Value': pfaId, 'Status': '', 'Notes': '' }),
  ];
  if (adjustmentsText) newRows.push(Object.assign({}, context, { 'Verification Item': 'Adjustments or Corrections completed if applicable', 'Value Type': 'Text', 'Unit': '', 'Actual Value': adjustmentsText, 'Status': '', 'Notes': '' }));
  appendObjectsAsRows_(getSuLogSheet_(), newRows);
  qualifyRun_(detail.runId);
  return { pfaId: pfaId };
}
function signPfa(recordId, qcName, adjustmentsText) { return signPfa_(recordId, qcName, adjustmentsText); }
