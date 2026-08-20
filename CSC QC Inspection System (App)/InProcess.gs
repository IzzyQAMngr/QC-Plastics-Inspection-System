/*************************************************************
 * PLASTICS IN-PROCESS INSPECTION — web app backend
 * Ported from legacy Plastics In-Process Inspection/Code.js
 * (saveQcRecord_ / getSpecsFromMaster_ / updateRowStatus_ / updateWallThickness_)
 *************************************************************/

// Chime A/B are each measured at two points (min & max reading) but the register carries
// only one LSL/USL band per chime — both readings are checked against that same band.
const CHIME_CHARACTERISTICS = ['Chime A', 'Chime B'];
// Exact register spelling for the 4 sidewall points that feed the wall-thickness-variance check.
const WALL_SW_CHARACTERISTICS = [
  'SW Thickness 1 - Top', 'SW Thickness 2 - Bottom', 'SW Thickness 3 - Side Right', 'SW Thickness 4 - Side Left',
];

/** Builds the mold's dimensional field list straight from its Spec Matrix rows — whatever
 *  characteristics the register defines for this mold, nothing hardcoded/universal.
 *  rejectLsl/rejectUsl (derived from the register's optional Reject Limit %) define a soft band
 *  just outside LSL/USL — an excursion inside that band is "Needs Review", not an automatic Fail.
 *  Blank Reject Limit % (the common case today) means rejectLsl/rejectUsl are both null, and
 *  evalSpec_ falls back to the original hard LSL/USL behavior exactly as before. */
function buildDimFields_(specs) {
  const fields = [];
  specs.forEach(spec => {
    const rejectLsl = (spec.lsl !== null && spec.rejectLimitPct !== null) ? spec.lsl * (1 - spec.rejectLimitPct / 100) : null;
    const rejectUsl = (spec.usl !== null && spec.rejectLimitPct !== null) ? spec.usl * (1 + spec.rejectLimitPct / 100) : null;
    if (CHIME_CHARACTERISTICS.indexOf(spec.characteristic) >= 0) {
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
  // Display order follows the Spec Matrix row order for this mold exactly — no code-side
  // reordering. Keep each mold's register rows arranged in the sequence you want on the form.
  return fields;
}

function getInProcessLogSheet_() {
  const sheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + INPROCESS_LOG_SHEET_NAME + '" sheet not found.');
  return sheet;
}

/** Called by InProcessView.html on load. Mold-first: no more Product-Type pre-filter — picking
 *  a Mold auto-fills its Product Type + Description client-side from this same list. */
function getInProcessFormData() {
  return {
    molds: getAllMoldsList_(),        // [{moldId, description, productType}]
    itemList: getItemList_(),         // [{itemNo, description}]
    inspectors: getInspectorList_(),
    foremen: getForemanList_(),
    shifts: getShiftList_(),
    passFailOptions: getPassFailNAList_(),
    releaseDecisionOptions: RELEASE_DECISION_OPTIONS,
    runs: getActiveRuns_(),
  };
}

function getColorOptions(mold) { return getColorOptionsForMold_(mold); }
function getCavityOptions(mold) { return getCavityIds_(mold); }
/** Re-fetched whenever Color or Item No. changes — colorSpec depends on both, not just the mold. */
function getColorSpecForRow(mold, color, itemNo) { return color ? getColorSpec_(mold, color, itemNo) : null; }

/** Returns this mold's dimensional field list + LSL/USL bounds + color spec, for the client to render/highlight. */
function getSpecsForRow(mold, color, itemNo) {
  const fields = buildDimFields_(getSpecsFromMaster_(mold));
  const bounds = {};
  fields.forEach(f => { bounds[f.key] = { lsl: f.lsl, nominal: f.nominal, usl: f.usl, rejectLsl: f.rejectLsl, rejectUsl: f.rejectUsl, measureIndex: f.measureIndex }; });
  const colorSpec = color ? getColorSpec_(mold, color, itemNo) : null;
  return { fields: fields, bounds: bounds, colorSpec: colorSpec, cavityIds: getCavityIds_(mold) };
}

/** 3-tier: Fail (beyond LSL/USL with no Reject Limit band, or beyond the band too), NeedsReview
 *  (past LSL/USL but still inside the register's Reject Limit % band), Pass. */
function evalSpec_(v, lsl, usl, rejectLsl, rejectUsl) {
  if (v === '' || v === null || v === undefined || isNaN(parseFloat(v))) return { status: '', detail: '' };
  const val = parseFloat(v);
  if (lsl !== null && !isNaN(lsl) && val < lsl) {
    if (rejectLsl !== null && !isNaN(rejectLsl) && val >= rejectLsl) return { status: 'NeedsReview', detail: val + ' < LSL ' + lsl };
    return { status: 'Fail', detail: val + ' < LSL ' + lsl };
  }
  if (usl !== null && !isNaN(usl) && val > usl) {
    if (rejectUsl !== null && !isNaN(rejectUsl) && val <= rejectUsl) return { status: 'NeedsReview', detail: val + ' > USL ' + usl };
    return { status: 'Fail', detail: val + ' > USL ' + usl };
  }
  return { status: 'Pass', detail: '' };
}

function displayStatus_(s) { return s === 'NeedsReview' ? 'Needs Review' : s; }

/** Server-side authoritative evaluation of one physical-sample row — mirrors updateRowStatus_/updateWallThickness_.
 *  `fields` is this mold's dynamic dimensional field list (see buildDimFields_), not a universal set. */
function evaluateInProcessRow_(row, specs) {
  const fields = buildDimFields_(specs);
  const hasAllWallFields = WALL_SW_CHARACTERISTICS.every(c => fields.some(f => f.key === c));

  let swVar = '', swEval = '';
  if (hasAllWallFields) {
    const wallVals = WALL_SW_CHARACTERISTICS.map(k => parseFloat(row[k]));
    if (wallVals.every(v => !isNaN(v))) {
      swVar = Math.round((Math.max(...wallVals) - Math.min(...wallVals)) * 10000) / 10000;
      swEval = swVar < 0.005 ? 'Pass' : 'Fail';
    }
  }

  const failures = [];
  const needsReview = [];
  const dateCode = String(row.dateCode || '').toLowerCase();
  if (dateCode === 'needs update') failures.push('Date Code Needs Update');

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

  if (swEval.toLowerCase() === 'fail') failures.push('SW Var. ' + swVar + ' ≥ 0.005');

  // ΔE*ab vs the register's Color Specs ΔE* Max — a magnitude, so USL-only (no LSL).
  // Only L*/a*/b* deltas are recorded on the form; the target L*/a*/b* themselves aren't shown or used.
  let colorSpec = null, deltaEMax = null, deltaEStatus = '';
  if (row.color) {
    try { colorSpec = getColorSpec_(row.mold, row.color, row.itemNo); } catch (e) { colorSpec = null; }
    const max = colorSpec ? parseFloat(colorSpec.deltaEMax) : NaN;
    if (!isNaN(max)) {
      deltaEMax = max;
      const de = parseFloat(row.deltaE);
      if (!isNaN(de)) {
        deltaEStatus = de > max ? 'Fail' : 'Pass';
        if (deltaEStatus === 'Fail') failures.push('ΔE*ab ' + de + ' > ' + max);
      }
    }
  }

  [['nesting', 'Nesting'], ['coverFit', 'Cover Fit'], ['gaugeFit', 'Gauge Fit']].forEach(([key, label]) => {
    if (String(row[key] || '').toLowerCase() === 'fail') failures.push(label + ' Fail');
  });

  const hasData = fields.some(f => dimResults[f.key] !== null) || dateCode !== '' ||
    row.nesting || row.coverFit || row.gaugeFit;

  // Needs Review never escalates to Fail on its own — only a genuine Fail (beyond any Reject
  // Limit band, or a characteristic with no band at all) triggers the auto-fail email below.
  let overallStatus = '';
  if (failures.length > 0) overallStatus = 'FAIL: ' + failures.join('; ');
  else if (needsReview.length > 0) overallStatus = 'NEEDS REVIEW: ' + needsReview.join('; ');
  else if (hasData) overallStatus = 'PASS';

  return {
    swVar: swVar, swEval: swEval, dimResults: dimResults, overallStatus: overallStatus,
    failures: failures, needsReview: needsReview, fields: fields,
    colorSpec: colorSpec, deltaEMax: deltaEMax, deltaEStatus: deltaEStatus,
  };
}

function makeRecordID_(sheet) {
  return makeSequentialId_(sheet, 'QC Record #', 'QC');
}

/**
 * Saves a full In-Process inspection (one or more physical-sample rows).
 * payload: { recordId (nullable), header: {inspDate,inspTime,shift,inspector,foreman},
 *   rows: [{runId,line,product,mold,color,resinLot,pallet,sampleDate,sampleTime,cavity,visual,
 *     dateCode,deltaL,deltaA,deltaB,deltaE,nesting,coverFit,gaugeFit,
 *     ...dynamic per-mold dimensional keys from buildDimFields_}] }
 */
function saveInProcessInspection(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const db = getInProcessLogSheet_();
    const header = payload.header || {};
    const rows = (payload.rows || []).filter(r => String(r.mold || '').trim());
    if (rows.length === 0) throw new Error('No inspection rows with a Mold selected.');

    let recordID = String(payload.recordId || '').trim();
    if (!/^QC-\d{6}-\d+$/.test(recordID)) recordID = makeRecordID_(db);

    const timestamp = new Date();
    // Month/Year reflect the Inspection Date (matches the sheet's old MONTH()/YEAR() formulas —
    // and what you'd actually want for monthly/yearly rollups), not the moment the save button
    // was clicked. Parsed by splitting the "yyyy-mm-dd" <input type=date> string directly rather
    // than `new Date(header.inspDate)`, which would shift by a day around UTC midnight in some
    // timezones. Falls back to the save timestamp if Inspection Date is missing/malformed.
    const inspDateParts = String(header.inspDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const month = inspDateParts ? Number(inspDateParts[2]) : timestamp.getMonth() + 1;
    const year = inspDateParts ? Number(inspDateParts[1]) : timestamp.getFullYear();
    const specCache = {};
    const getSpecs = (mold) => {
      if (!specCache[mold]) {
        try { specCache[mold] = getSpecsFromMaster_(mold); } catch (e) { specCache[mold] = []; }
      }
      return specCache[mold];
    };

    // Evaluate every row up front — a deviating cavity (server-side FAIL) missing its Release
    // Decision + Justification blocks the entire save, never a partial write. Mirrors the
    // client-side gate in InProcessView.html's renderReview(), which is what a QC tech actually
    // sees; this is the authoritative backstop in case that client check is ever bypassed.
    const evalResults = rows.map(row => evaluateInProcessRow_(row, getSpecs(row.mold)));
    const unresolved = [];
    rows.forEach((row, i) => {
      if (evalResults[i].overallStatus.toUpperCase().indexOf('FAIL') === 0) {
        if (!String(row.releaseDecision || '').trim() || !String(row.releaseJustification || '').trim()) {
          unresolved.push('Line ' + row.line + ' Cavity ' + (row.cavity || '—'));
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
          'Inspection Date': header.inspDate, 'Inspection Time': header.inspTime, 'Inspected By': header.inspector,
          'Shift': header.shift, 'Shift Foreman': header.foreman, 'Line #': row.line, 'Product Type': row.product,
          'Mold': row.mold, 'Color': row.color, 'LOT of Resin': row.resinLot, 'Pallet Sequence': row.pallet,
          'Sample Date': row.sampleDate, 'Sample Time': row.sampleTime, 'Cavity ID': row.cavity,
          'Test Type': testType, 'Measure Index': measIdx, 'Characteristic Name': charName, 'Unit': unit,
          'LSL': lsl, 'USL': usl, 'Actual Value': actual, 'Status': status, 'Status Detail': detail,
          'Visual Notes': row.visual, 'Source': 'CSC QC Inspection System', 'Month': month, 'Year': year,
          'BatchID': row.runId || payload.runId || '',
          'Item No.': row.itemNo || '', 'Item Description': row.itemDescription || '',
          'Release Decision': row.releaseDecision || '', 'Justification': row.releaseJustification || '',
        };
      }

      if (row.dateCode) {
        const dc = String(row.dateCode).toLowerCase();
        const dcStatus = dc === 'needs update' ? 'Fail' : (dc === 'current' ? 'Pass' : '');
        dbRows.push(dbRow('Visual', '', 'Date Code Verification', '', '', '', row.dateCode, dcStatus,
          dcStatus === 'Fail' ? 'Date code needs update' : ''));
      }
      if (row.visual) dbRows.push(dbRow('Visual', '', 'Visual Observations', '', '', '', row.visual, '', ''));

      [['deltaL', 'ΔL'], ['deltaA', 'ΔA'], ['deltaB', 'ΔB']].forEach(([key, name]) => {
        if (row[key]) dbRows.push(dbRow('Color', '', name, '', '', '', row[key], '', ''));
      });
      if (row.deltaE) {
        dbRows.push(dbRow('Color', '', 'ΔE*ab', '', '', evalResult.deltaEMax !== null ? evalResult.deltaEMax : '',
          row.deltaE, evalResult.deltaEStatus, evalResult.deltaEStatus === 'Fail' ? 'ΔE*ab exceeds max' : ''));
      }

      evalResult.fields.forEach(f => {
        const v = row[f.key];
        if (!v && v !== 0) return;
        const r = evalResult.dimResults[f.key];
        dbRows.push(dbRow('Dimensional', f.measureIndex, f.key, f.unit, r ? r.lsl : '', r ? r.usl : '', v,
          r ? displayStatus_(r.status) : '', r && r.status === 'NeedsReview' ? 'Within reject limit' : ''));
      });
      if (evalResult.swVar !== '') {
        dbRows.push(dbRow('Dimensional', '', 'Wall Thickness Variance', 'in', '', '', evalResult.swVar, '', ''));
        dbRows.push(dbRow('Dimensional', '', 'Wall Thickness Evaluation', '', '', '', evalResult.swEval, '', ''));
      }

      [['nesting', 'Nesting'], ['coverFit', 'Cover Fit'], ['gaugeFit', 'Gauge Fit']].forEach(([key, name]) => {
        const v = row[key];
        if (!v) return;
        const s = String(v).toLowerCase();
        const st = (s === 'pass' || s === 'good' || s === 'x') ? 'Pass' : ((s === 'fail' || s === 'loose' || s === 'tight') ? 'Fail' : '');
        dbRows.push(dbRow('Functional', '', name, '', '', '', v, st, st === 'Fail' ? name + ' Fail' : ''));
      });

      if (evalResult.overallStatus.toUpperCase().indexOf('FAIL') === 0) {
        failRows.push({
          mold: row.mold, cavity: row.cavity, status: evalResult.overallStatus, product: row.product, line: row.line,
          releaseDecision: row.releaseDecision || '', releaseJustification: row.releaseJustification || '',
        });
      } else if (evalResult.overallStatus.toUpperCase().indexOf('NEEDS REVIEW') === 0) {
        reviewCount++;
      }
    });

    if (dbRows.length === 0) throw new Error('No data to save.');
    appendObjectsAsRows_(db, dbRows);

    if (failRows.length > 0) {
      try { sendInProcessFailNotification_(recordID, header, failRows); } catch (e) { /* best-effort */ }
    }

    return { recordId: recordID, savedRows: dbRows.length, failCount: failRows.length, reviewCount: reviewCount };
  } finally {
    lock.releaseLock();
  }
}

function sendInProcessFailNotification_(recordID, header, failRows) {
  const emails = getNotificationEmails_();
  if (emails.length === 0) return;

  let html = '<div style="font-family:Calibri,sans-serif;max-width:700px;">';
  html += '<div style="background:#1A3A3A;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">';
  html += '<h2 style="margin:0;font-size:18px;">⚠️ QC DEVIATION NOTIFICATION</h2>';
  html += '<p style="margin:4px 0 0;font-size:13px;opacity:0.8;">CSC QC Inspection System — Plastics In-Process</p></div>';
  html += '<div style="background:#f8f9fa;padding:16px 20px;border:1px solid #dee2e6;">';
  html += '<table style="font-size:13px;margin-bottom:12px;">';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">QC Record #:</td><td><b>' + recordID + '</b></td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Date:</td><td>' + header.inspDate + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Inspector:</td><td>' + header.inspector + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Shift:</td><td>' + header.shift + '</td></tr>';
  html += '</table>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
  html += '<tr style="background:#d9534f;color:white;"><th style="padding:8px;text-align:left;">Line</th>' +
    '<th style="padding:8px;text-align:left;">Mold</th><th style="padding:8px;text-align:left;">Cavity</th>' +
    '<th style="padding:8px;text-align:left;">Deviation Details</th>' +
    '<th style="padding:8px;text-align:left;">Release Decision</th><th style="padding:8px;text-align:left;">Justification</th></tr>';
  failRows.forEach((f, i) => {
    const bg = i % 2 === 0 ? '#fff' : '#f8f8f8';
    const detail = f.status.replace(/^FAIL:\s*/i, '');
    html += '<tr style="background:' + bg + ';"><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.line || '') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.mold +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.cavity +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;color:#d9534f;font-weight:bold;">' + detail +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.releaseDecision || '—') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.releaseJustification || '—') + '</td></tr>';
  });
  html += '</table></div>';
  html += '<div style="background:#eee;padding:10px 20px;border-radius:0 0 8px 8px;font-size:11px;color:#888;">' +
    'Sent automatically by the CSC QC Inspection System.</div></div>';

  const moldList = [...new Set(failRows.map(f => f.mold))].join(', ');
  const subject = '⚠️ QC PLASTICS INSP. DEVIATION — ' + recordID + ' | ' + moldList;
  emails.forEach(to => MailApp.sendEmail({ to: to, subject: subject, htmlBody: html }));
}
