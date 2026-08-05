/*************************************************************
 * PLASTICS IN-PROCESS INSPECTION — web app backend
 * Ported from legacy Plastics In-Process Inspection/Code.js
 * (saveQcRecord_ / getSpecsFromMaster_ / updateRowStatus_ / updateWallThickness_)
 *************************************************************/

const DIM_DEFS = [
  { key: 'weight',    name: 'Weight',                      unit: 'g',  short: 'Wt',     match: ['weight'] },
  { key: 'botThick',  name: 'Bottom Thickness',            unit: 'in', short: 'BotTh',  match: ['bottom thickness'] },
  { key: 'sw1',       name: 'SW Thickness 1 - Top',        unit: 'in', short: 'SW1',    match: ['sw thickness 1', 'side wall thickness 1'] },
  { key: 'sw2',       name: 'SW Thickness 2 - Bottom',     unit: 'in', short: 'SW2',    match: ['sw thickness 2', 'side wall thickness 2'] },
  { key: 'sw3',       name: 'SW Thickness 3 - Side Right', unit: 'in', short: 'SW3',    match: ['sw thickness 3', 'side wall thickness 3'] },
  { key: 'sw4',       name: 'SW Thickness 4 - Side Left',  unit: 'in', short: 'SW4',    match: ['sw thickness 4', 'side wall thickness 4'] },
  { key: 'topDia',    name: 'Top Diameter',                unit: 'in', short: 'TopDia', match: ['top diameter', 'top od'] },
  { key: 'botDia',    name: 'Bottom Diameter',             unit: 'in', short: 'BotDia', match: ['bottom diameter', 'bottom od'] },
  { key: 'endPin',    name: 'End Pin Diameter',            unit: 'in', short: 'EndPin', match: ['end pin'] },
  { key: 'height',    name: 'Height',                      unit: 'in', short: 'Ht',     match: ['height'] },
  { key: 'chimeAMin', name: 'Chime A Min',                 unit: 'in', short: 'ChA.min', match: ['chime a'] },
  { key: 'chimeAMax', name: 'Chime A Max',                 unit: 'in', short: 'ChA.max', match: ['chime a'] },
  { key: 'chimeBMin', name: 'Chime B Min',                 unit: 'in', short: 'ChB.min', match: ['chime b'] },
  { key: 'chimeBMax', name: 'Chime B Max',                 unit: 'in', short: 'ChB.max', match: ['chime b'] },
];
const WALL_KEYS = ['sw1', 'sw2', 'sw3', 'sw4'];

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
    openBatches: getOpenBatches_(),
  };
}

function getColorOptions(mold) { return getColorOptionsForMold_(mold); }
function getCavityOptions(mold) { return getCavityIds_(mold); }

/** Returns spec info (LSL/USL per dimension + color spec) for a Mold/Color/Item combo, for live client-side highlighting. */
function getSpecsForRow(productType, mold, color, itemNo) {
  const specs = getSpecsFromMaster_(productType, mold);
  const bounds = {};
  DIM_DEFS.forEach(d => {
    const spec = findSpecForMatch_(specs, d.match);
    bounds[d.key] = spec ? { lsl: spec.lsl, usl: spec.usl, measureIndex: spec.measureIndex } : null;
  });
  const colorSpec = color ? getColorSpec_(mold, color, itemNo) : null;
  return { bounds: bounds, colorSpec: colorSpec, cavityIds: getCavityIds_(mold) };
}

function findSpecForMatch_(specs, matchTerms) {
  for (const spec of specs) {
    const name = String(spec.characteristic).toLowerCase()
      .replace(/\(.*?\)/g, '').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
    for (const term of matchTerms) {
      if (name === term || name.indexOf(term) === 0 || term.indexOf(name) === 0) return spec;
    }
  }
  return null;
}

function evalSpec_(v, lsl, usl) {
  if (v === '' || v === null || v === undefined || isNaN(parseFloat(v))) return { status: '', detail: '' };
  const val = parseFloat(v);
  if (lsl !== null && !isNaN(lsl) && val < lsl) return { status: 'Fail', detail: val + ' < LSL ' + lsl };
  if (usl !== null && !isNaN(usl) && val > usl) return { status: 'Fail', detail: val + ' > USL ' + usl };
  return { status: 'Pass', detail: '' };
}

/** Server-side authoritative evaluation of one physical-sample row — mirrors updateRowStatus_/updateWallThickness_. */
function evaluateInProcessRow_(row, specs) {
  const bounds = {};
  DIM_DEFS.forEach(d => { bounds[d.key] = findSpecForMatch_(specs, d.match); });

  const wallVals = WALL_KEYS.map(k => row[k]).map(v => parseFloat(v));
  let swVar = '', swEval = '';
  if (wallVals.every(v => !isNaN(v))) {
    swVar = Math.round((Math.max(...wallVals) - Math.min(...wallVals)) * 10000) / 10000;
    swEval = swVar < 0.005 ? 'Pass' : 'Fail';
  }

  const failures = [];
  const dateCode = String(row.dateCode || '').toLowerCase();
  if (dateCode === 'fail') failures.push('Date Code Fail');

  const dimResults = {};
  DIM_DEFS.forEach(d => {
    const spec = bounds[d.key];
    const raw = row[d.key];
    if (raw === '' || raw === null || raw === undefined || isNaN(parseFloat(raw))) { dimResults[d.key] = null; return; }
    const value = parseFloat(raw);
    const lsl = spec ? spec.lsl : null, usl = spec ? spec.usl : null;
    const ev = evalSpec_(value, lsl, usl);
    dimResults[d.key] = { value: value, lsl: lsl, usl: usl, status: ev.status };
    if (lsl !== null && !isNaN(lsl) && value < lsl) failures.push(d.short + ' ' + value + ' < LSL ' + lsl);
    if (usl !== null && !isNaN(usl) && value > usl) failures.push(d.short + ' ' + value + ' > USL ' + usl);
  });

  if (swEval.toLowerCase() === 'fail') failures.push('SW Var. ' + swVar + ' ≥ 0.005');

  [['nesting', 'Nesting'], ['coverFit', 'Cover Fit'], ['gaugeFit', 'Gauge Fit']].forEach(([key, label]) => {
    if (String(row[key] || '').toLowerCase() === 'fail') failures.push(label + ' Fail');
  });

  const hasData = DIM_DEFS.some(d => dimResults[d.key] !== null) || dateCode !== '' ||
    row.nesting || row.coverFit || row.gaugeFit;

  let overallStatus = '';
  if (failures.length > 0) overallStatus = 'FAIL: ' + failures.join('; ');
  else if (hasData) overallStatus = 'PASS';

  return { swVar: swVar, swEval: swEval, dimResults: dimResults, overallStatus: overallStatus, failures: failures };
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
 *   rows: [{batchId,line,product,mold,color,resinLot,pallet,sampleDate,sampleTime,cavity,visual,
 *     dateCode,deltaL,deltaA,deltaB,deltaE,weight,botThick,sw1,sw2,sw3,sw4,topDia,botDia,endPin,
 *     height,chimeAMin,chimeAMax,chimeBMin,chimeBMax,nesting,coverFit,gaugeFit}] }
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
    const getSpecs = (product, mold) => {
      const key = product + '|' + mold;
      if (!specCache[key]) {
        try { specCache[key] = getSpecsFromMaster_(product, mold); } catch (e) { specCache[key] = []; }
      }
      return specCache[key];
    };

    const dbRows = [];
    const failRows = [];
    let inspectionId = 0;

    rows.forEach(row => {
      inspectionId++;
      const specs = getSpecs(row.product, row.mold);
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
          'BatchID': row.batchId || payload.batchId || '',
          'Item No.': row.itemNo || '', 'Item Description': row.itemDescription || '',
        };
      }

      if (row.dateCode) {
        const dc = String(row.dateCode).toLowerCase();
        const dcStatus = dc === 'fail' ? 'Fail' : (dc === 'pass' ? 'Pass' : '');
        dbRows.push(dbRow('Visual', '', 'Date Code Verification', '', '', '', row.dateCode, dcStatus,
          dcStatus === 'Fail' ? 'Date code incorrect' : ''));
      }
      if (row.visual) dbRows.push(dbRow('Visual', '', 'Visual Conformance Check', '', '', '', row.visual, '', ''));

      [['deltaL', 'ΔL'], ['deltaA', 'ΔA'], ['deltaB', 'ΔB'], ['deltaE', 'ΔE*ab']].forEach(([key, name]) => {
        if (row[key]) dbRows.push(dbRow('Color', '', name, '', '', '', row[key], '', ''));
      });

      DIM_DEFS.forEach(d => {
        const v = row[d.key];
        if (!v && v !== 0) return;
        const r = evalResult.dimResults[d.key];
        dbRows.push(dbRow('Dimensional', findSpecForMatch_(specs, d.match) ? findSpecForMatch_(specs, d.match).measureIndex : '',
          d.name, d.unit, r ? r.lsl : '', r ? r.usl : '', v, r ? r.status : '', ''));
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
      }
    });

    if (dbRows.length === 0) throw new Error('No data to save.');
    appendObjectsAsRows_(db, dbRows);

    if (failRows.length > 0) {
      try { sendInProcessFailNotification_(recordID, header, failRows); } catch (e) { /* best-effort */ }
    }

    return { recordId: recordID, savedRows: dbRows.length, failCount: failRows.length };
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
