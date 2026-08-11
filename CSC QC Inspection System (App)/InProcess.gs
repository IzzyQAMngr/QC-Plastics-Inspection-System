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
  return sortDimFields_(fields);
}

// Display order requested for the Dimensional section: Weight, Pail Height, Chime A/B Min/Max
// first (in that order); any other register-defined characteristic keeps its original relative
// order in the middle; Wall Thickness (the 4 SW sidewall points) always sorts last.
const DIM_FIELD_PRIORITY_ = ['Weight', 'Pail Height', 'Chime A Min', 'Chime A Max', 'Chime B Min', 'Chime B Max'];
function sortDimFields_(fields) {
  const rank = f => {
    const p = DIM_FIELD_PRIORITY_.indexOf(f.key);
    if (p >= 0) return p;
    if (WALL_SW_CHARACTERISTICS.indexOf(f.key) >= 0) return 1000 + WALL_SW_CHARACTERISTICS.indexOf(f.key);
    return 100;
  };
  return fields.map((f, i) => ({ f: f, i: i })).sort((a, b) => (rank(a.f) - rank(b.f)) || (a.i - b.i)).map(x => x.f);
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
  const tz = getDb_().getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const rows = readSheetObjects_(sheet);
  let max = 0;
  rows.forEach(r => {
    const m = String(r['QC Record #'] || '').match(/^QC-(\d{8})-(\d{3})$/);
    if (m && m[1] === today) { const s = Number(m[2]); if (s > max) max = s; }
  });
  return 'QC-' + today + '-' + String(max + 1).padStart(3, '0');
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
    if (!/^QC-\d{8}-\d{3}$/.test(recordID)) recordID = makeRecordID_(db);

    const timestamp = new Date();
    const month = timestamp.getMonth() + 1;
    const year = timestamp.getFullYear();
    const specCache = {};
    const getSpecs = (mold) => {
      if (!specCache[mold]) {
        try { specCache[mold] = getSpecsFromMaster_(mold); } catch (e) { specCache[mold] = []; }
      }
      return specCache[mold];
    };

    const dbRows = [];
    const failRows = [];
    let reviewCount = 0;
    let inspectionId = 0;

    rows.forEach(row => {
      inspectionId++;
      const specs = getSpecs(row.mold);
      const evalResult = evaluateInProcessRow_(row, specs);

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
        failRows.push({ mold: row.mold, cavity: row.cavity, status: evalResult.overallStatus, product: row.product, line: row.line });
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
  html += '<h2 style="margin:0;font-size:18px;">⚠️ QC FAIL NOTIFICATION</h2>';
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
    '<th style="padding:8px;text-align:left;">Failure Details</th></tr>';
  failRows.forEach((f, i) => {
    const bg = i % 2 === 0 ? '#fff' : '#f8f8f8';
    const detail = f.status.replace(/^FAIL:\s*/i, '');
    html += '<tr style="background:' + bg + ';"><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.line || '') +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.mold +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.cavity +
      '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;color:#d9534f;font-weight:bold;">' + detail + '</td></tr>';
  });
  html += '</table></div>';
  html += '<div style="background:#eee;padding:10px 20px;border-radius:0 0 8px 8px;font-size:11px;color:#888;">' +
    'Sent automatically by the CSC QC Inspection System.</div></div>';

  const moldList = [...new Set(failRows.map(f => f.mold))].join(', ');
  const subject = '⚠️ QC PLASTICS INSP. FAIL — ' + recordID + ' | ' + moldList;
  emails.forEach(to => MailApp.sendEmail({ to: to, subject: subject, htmlBody: html }));
}
