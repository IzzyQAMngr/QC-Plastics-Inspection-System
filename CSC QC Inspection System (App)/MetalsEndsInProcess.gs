/*************************************************************
 * METALS ENDS — IN-PROCESS INSPECTION — web app backend
 * Mirrors Plastics In-Process (InProcess.gs) — spec-driven dimensional fields pulled live from
 * the Metals Spec Register's Spec Matrix (Product Type = 'End'), Line-tabs + repeatable sample
 * rows per Line, Release Decision/Justification gate on any deviation, auto-fail email.
 *
 * Two things Plastics In-Process doesn't have, both by explicit request (2026-08-13):
 * - Material Verification (Tin Plate Tag # / Basis Wt / Temper / Coating: Tag vs Actual) is
 *   recorded as-is with no automated Tag/Actual comparison — purely the tech's own judgment call.
 * - Curl Diameter / Countersink / Curl Thickness each get a Min + Max reading checked against
 *   the SAME spec band, exactly like Plastics' Chime A/B — see MIN_MAX_CHARACTERISTICS_ENDS.
 *************************************************************/

const MIN_MAX_CHARACTERISTICS_ENDS = ['Curl Diameter', 'Countersink', 'Curl Thickness'];
const FUNCTIONAL_TEST_KEYS_ENDS = [['pingGauge', 'Ping Gauge'], ['copperSulfate', 'Copper Sulfate'], ['scotchTape', 'Scotch Tape']];
const MATERIAL_VERIFICATION_KEYS_ENDS = [
  ['tinPlateTag', 'Tin Plate Tag #'], ['basisWtTag', 'Basis Wt (Tag)'], ['basisWtActual', 'Basis Wt (Actual)'],
  ['temperTag', 'Temper (Tag)'], ['temperActual', 'Temper (Actual)'], ['coatingTag', 'Coating (Tag)'], ['coatingActual', 'Coating (Actual)'],
];

/** Builds this End Size's dimensional field list from its Spec Matrix rows — Curl Diameter /
 *  Countersink / Curl Thickness each become two fields (Min + Max) sharing one spec band,
 *  everything else (Basis Weight, 2 Inch Stack Count, Compound Film Weight) stays a single field. */
function buildEndDimFields_(specs) {
  const fields = [];
  specs.forEach(spec => {
    const rejectLsl = (spec.lsl !== null && spec.rejectLimitPct !== null) ? spec.lsl * (1 - spec.rejectLimitPct / 100) : null;
    const rejectUsl = (spec.usl !== null && spec.rejectLimitPct !== null) ? spec.usl * (1 + spec.rejectLimitPct / 100) : null;
    if (MIN_MAX_CHARACTERISTICS_ENDS.indexOf(spec.characteristic) >= 0) {
      ['Min', 'Max'].forEach(suffix => {
        const key = spec.characteristic + ' ' + suffix;
        fields.push({
          key: key, characteristic: spec.characteristic, label: key + (spec.unit ? ' (' + spec.unit + ')' : ''),
          unit: spec.unit, lsl: spec.lsl, nominal: spec.nominal, usl: spec.usl, rejectLsl: rejectLsl, rejectUsl: rejectUsl, measureIndex: spec.measureIndex,
        });
      });
    } else {
      fields.push({
        key: spec.characteristic, characteristic: spec.characteristic,
        label: spec.characteristic + (spec.unit ? ' (' + spec.unit + ')' : ''),
        unit: spec.unit, lsl: spec.lsl, nominal: spec.nominal, usl: spec.usl, rejectLsl: rejectLsl, rejectUsl: rejectUsl, measureIndex: spec.measureIndex,
      });
    }
  });
  return fields;
}

function getMetalsEndsLogSheet_() {
  const sheet = getDb_().getSheetByName(METALS_ENDS_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + METALS_ENDS_LOG_SHEET_NAME + '" sheet not found — run setupMetalsEndsLogSheet from the script editor first.');
  return sheet;
}

/** Called by MetalsEndsInProcessView.html on load. */
function getMetalsEndsFormData() {
  return {
    endSizes: getEndSizeList_(),           // [{sizeId, customer, description}]
    lines: getLinesForDepartment_('Metals'),
    inspectors: getInspectorList_('Metals'),
    foremen: getForemanList_('Metals'),
    shifts: getShiftList_('Metals'),
    functionalTestOptions: ['Pass', 'Fail', 'N/A'],
    releaseDecisionOptions: RELEASE_DECISION_OPTIONS,
  };
}

/** Machine ID options for one Line — re-fetched whenever the active Line tab changes. */
function getMachineIdsForEndsLine(line) { return getMachineIdsForLine_('Metals', line); }

/** This End Size's dimensional field list + LSL/USL bounds, for the client to render/highlight. */
function getSpecsForEndSizeRow(sizeId) {
  const fields = buildEndDimFields_(getSpecsForEndSize_(sizeId));
  const bounds = {};
  fields.forEach(f => { bounds[f.key] = { lsl: f.lsl, nominal: f.nominal, usl: f.usl, rejectLsl: f.rejectLsl, rejectUsl: f.rejectUsl, measureIndex: f.measureIndex }; });
  return { fields: fields, bounds: bounds };
}

/** Server-side authoritative evaluation of one sample row — mirrors evaluateInProcessRow_.
 *  Only Dimensional readings and Functional Test results can produce a Fail; Material
 *  Verification is recorded but never evaluated (Izzy's call, 2026-08-13 — "not completely sure
 *  how this is going to work" yet). */
function evaluateEndsRow_(row, specs) {
  const fields = buildEndDimFields_(specs);
  const failures = [];
  const needsReview = [];

  const dimResults = {};
  fields.forEach(f => {
    const raw = row[f.key];
    if (raw === '' || raw === null || raw === undefined || isNaN(parseFloat(raw))) { dimResults[f.key] = null; return; }
    const value = parseFloat(raw);
    const ev = evalSpec_(value, f.lsl, f.usl, f.rejectLsl, f.rejectUsl);
    dimResults[f.key] = { value: value, lsl: f.lsl, usl: f.usl, status: ev.status };
    if (ev.status === 'Fail') failures.push(f.key + ' ' + ev.detail);
    else if (ev.status === 'NeedsReview') needsReview.push(f.key + ' ' + ev.detail);
  });

  FUNCTIONAL_TEST_KEYS_ENDS.forEach(([key, label]) => {
    if (String(row[key] || '').trim() === 'Fail') failures.push(label + ' Fail');
  });

  const hasData = fields.some(f => dimResults[f.key] !== null) ||
    FUNCTIONAL_TEST_KEYS_ENDS.some(([key]) => row[key]) || row.visualNotes ||
    MATERIAL_VERIFICATION_KEYS_ENDS.some(([key]) => row[key]);

  let overallStatus = '';
  if (failures.length > 0) overallStatus = 'FAIL: ' + failures.join('; ');
  else if (needsReview.length > 0) overallStatus = 'NEEDS REVIEW: ' + needsReview.join('; ');
  else if (hasData) overallStatus = 'PASS';

  return { dimResults: dimResults, overallStatus: overallStatus, failures: failures, needsReview: needsReview, fields: fields };
}

function makeEndsRecordID_(sheet) { return makeSequentialId_(sheet, 'QC Record #', 'QCE'); }

/**
 * Saves a full Metals Ends In-Process inspection (one or more sample rows).
 * payload: { recordId (nullable), header: {inspDate,inspTime,shift,inspector,foreman},
 *   rows: [{line,machineId,endSize,endCustomer,endDescription,inspectionTime,visualNotes,
 *     tinPlateTag,basisWtTag,basisWtActual,temperTag,temperActual,coatingTag,coatingActual,
 *     pingGauge,copperSulfate,scotchTape,releaseDecision,releaseJustification,
 *     ...dynamic per-End-Size dimensional keys from buildEndDimFields_}] }
 */
function saveMetalsEndsInspection(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const db = getMetalsEndsLogSheet_();
    const header = payload.header || {};
    const rows = (payload.rows || []).filter(r => String(r.endSize || '').trim());
    if (rows.length === 0) throw new Error('No inspection rows with an End Size selected.');

    let recordID = String(payload.recordId || '').trim();
    if (!/^QCE-\d{6}-\d+$/.test(recordID)) recordID = makeEndsRecordID_(db);

    const timestamp = new Date();
    const tz = getDb_().getSpreadsheetTimeZone();
    const inspDateParts = String(header.inspDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const month = inspDateParts ? Number(inspDateParts[2]) : timestamp.getMonth() + 1;
    const year = inspDateParts ? Number(inspDateParts[1]) : timestamp.getFullYear();

    const specCache = {};
    const getSpecs = (sizeId) => {
      if (!specCache[sizeId]) {
        try { specCache[sizeId] = getSpecsForEndSize_(sizeId); } catch (e) { specCache[sizeId] = []; }
      }
      return specCache[sizeId];
    };

    // Evaluate every row up front — a deviating sample (server-side FAIL) missing its Release
    // Decision + Justification blocks the entire save, mirroring saveInProcessInspection.
    const evalResults = rows.map(row => evaluateEndsRow_(row, getSpecs(row.endSize)));
    const unresolved = [];
    rows.forEach((row, i) => {
      if (evalResults[i].overallStatus.toUpperCase().indexOf('FAIL') === 0) {
        if (!String(row.releaseDecision || '').trim() || !String(row.releaseJustification || '').trim()) {
          unresolved.push('Line ' + row.line + ' — ' + (row.machineId || '—'));
        }
      }
    });
    if (unresolved.length > 0) {
      throw new Error('Release Decision + Justification required before saving: ' + unresolved.join(', ') + '.');
    }

    const dbRows = [];
    const failRows = [];
    let reviewCount = 0;
    let inspectionId = 0;

    rows.forEach((row, i) => {
      inspectionId++;
      const evalResult = evalResults[i];

      function dbRow(testType, measIdx, charName, unit, lsl, usl, actual, status, detail) {
        return {
          'QC Record #': recordID, 'Timestamp Saved': timestamp, 'Inspection ID': inspectionId,
          'Inspection Date': header.inspDate, 'Inspection Time': row.inspectionTime || header.inspTime, 'Inspected By': header.inspector,
          'Shift': header.shift, 'Shift Foreman': header.foreman, 'Line #': row.line, 'Machine ID': row.machineId || '',
          'End Size': row.endSize, 'Customer': row.endCustomer || '', 'End Description': row.endDescription || '',
          'Test Type': testType, 'Measure Index': measIdx, 'Characteristic Name': charName, 'Unit': unit,
          'LSL': lsl, 'USL': usl, 'Actual Value': (actual === undefined || actual === null) ? '' : actual,
          'Status': status, 'Status Detail': detail, 'Visual Notes': '', 'Source': 'CSC QC Inspection System',
          'Month': month, 'Year': year, 'Release Decision': row.releaseDecision || '', 'Justification': row.releaseJustification || '',
        };
      }

      if (row.visualNotes) dbRows.push(dbRow('Visual', '', 'Visual Conformance Check', '', '', '', row.visualNotes, '', ''));

      MATERIAL_VERIFICATION_KEYS_ENDS.forEach(([key, label]) => {
        const v = row[key];
        if (!v && v !== 0) return;
        dbRows.push(dbRow('Material Verification', '', label, '', '', '', v, '', ''));
      });

      evalResult.fields.forEach(f => {
        const v = row[f.key];
        if (!v && v !== 0) return;
        const r = evalResult.dimResults[f.key];
        dbRows.push(dbRow('Dimensional', f.measureIndex, f.key, f.unit, r ? r.lsl : '', r ? r.usl : '', v,
          r ? displayStatus_(r.status) : '', r && r.status === 'NeedsReview' ? 'Within reject limit' : ''));
      });

      FUNCTIONAL_TEST_KEYS_ENDS.forEach(([key, label]) => {
        const v = row[key];
        if (!v) return;
        dbRows.push(dbRow('Functional', '', label, '', '', '', v, v, v === 'Fail' ? label + ' Fail' : ''));
      });

      if (evalResult.overallStatus.toUpperCase().indexOf('FAIL') === 0) {
        failRows.push({
          endSize: row.endSize, machineId: row.machineId, line: row.line, status: evalResult.overallStatus,
          releaseDecision: row.releaseDecision || '', releaseJustification: row.releaseJustification || '',
        });
      } else if (evalResult.overallStatus.toUpperCase().indexOf('NEEDS REVIEW') === 0) {
        reviewCount++;
      }
    });

    if (dbRows.length === 0) throw new Error('No data to save.');
    appendObjectsAsRows_(db, dbRows);

    if (failRows.length > 0) {
      try { sendEndsFailNotification_(recordID, header, failRows); } catch (e) { /* best-effort */ }
    }

    return { recordId: recordID, savedRows: dbRows.length, failCount: failRows.length, reviewCount: reviewCount };
  } finally {
    lock.releaseLock();
  }
}

function sendEndsFailNotification_(recordID, header, failRows) {
  const emails = getNotificationEmails_();
  if (emails.length === 0) return;

  let html = '<div style="font-family:Calibri,sans-serif;max-width:700px;">';
  html += '<div style="background:#1A3A3A;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">';
  html += '<h2 style="margin:0;font-size:18px;">⚠️ QC DEVIATION NOTIFICATION</h2>';
  html += '<p style="margin:4px 0 0;font-size:13px;opacity:0.8;">CSC QC Inspection System — Metals Ends In-Process</p></div>';
  html += '<div style="background:#f8f9fa;padding:16px 20px;border:1px solid #dee2e6;">';
  html += '<table style="font-size:13px;margin-bottom:12px;">';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">QC Record #:</td><td><b>' + recordID + '</b></td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Date:</td><td>' + header.inspDate + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Inspector:</td><td>' + header.inspector + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Shift:</td><td>' + header.shift + '</td></tr>';
  html += '</table>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
  html += '<tr style="background:#d9534f;color:white;"><th style="padding:8px;text-align:left;">Line</th>' +
    '<th style="padding:8px;text-align:left;">Machine</th><th style="padding:8px;text-align:left;">End Size</th>' +
    '<th style="padding:8px;text-align:left;">Deviation Details</th>' +
    '<th style="padding:8px;text-align:left;">Release Decision</th><th style="padding:8px;text-align:left;">Justification</th></tr>';
  failRows.forEach((f, i) => {
    const bg = i % 2 === 0 ? '#fff' : '#f8f8f8';
    const detail = f.status.replace(/^FAIL:\s*/i, '');
    html += '<tr style="background:' + bg + ';"><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.line || '') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.machineId || '—') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.endSize +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;color:#d9534f;font-weight:bold;">' + detail +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.releaseDecision || '—') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.releaseJustification || '—') + '</td></tr>';
  });
  html += '</table></div>';
  html += '<div style="background:#eee;padding:10px 20px;border-radius:0 0 8px 8px;font-size:11px;color:#888;">' +
    'Sent automatically by the CSC QC Inspection System.</div></div>';

  const sizeList = [...new Set(failRows.map(f => f.endSize))].join(', ');
  const subject = '⚠️ QC METALS ENDS INSP. DEVIATION — ' + recordID + ' | ' + sizeList;
  emails.forEach(to => MailApp.sendEmail({ to: to, subject: subject, htmlBody: html }));
}
