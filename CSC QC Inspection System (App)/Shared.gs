/*************************************************************
 * CSC QC INSPECTION SYSTEM — SHARED CONFIG & HELPERS
 * Consolidated Settings / Master Register / Runs access
 * used by both Drop Freeze and In-Process modules.
 *************************************************************/

// ================= SHEET NAMES =================
const SETTINGS_SHEET_NAME       = 'Settings';
const DROPFREEZE_LOG_SHEET_NAME = 'Test Data';
const INPROCESS_LOG_SHEET_NAME  = 'QC Database';
const ALL_ITEMS_LIST_SHEET_NAME = 'All Items List'; // its own tab, not part of Settings
const SU_ITEMS_SHEET_NAME = 'Start-Up Verification Items List- Plastics';
const SU_LOG_SHEET_NAME   = 'Start-Up Verification Log- Plastics';

// Spec Register (relinked 2026-08 — Spec Matrix / Color Specs / All Molds List
// replaced the old per-product tabs). This is the default; Settings' own
// "Master Register ID" row can override it without a redeploy.
const SPEC_REGISTER_ID = '1rgm9gAnZviUSLKLF1kjbX0P1X5fPvGvKXDg5SvT1htM';

const ALL_MOLDS_LIST_SHEET = 'All Molds List';
const SPEC_MATRIX_SHEET    = 'Spec Matrix';
const COLOR_SPECS_SHEET    = 'Color Specs';
const FUNCTIONAL_TESTS_SHEET = 'Functional Tests';

// Clean header row for a brand-new Test Data (Drop Freeze) log — this working
// copy has no Drop Freeze data yet, so we're free to define it cleanly. Run-driven
// (2026-08-07): shares its context columns with the QC Database / Start-Up Log pattern.
const DROPFREEZE_LOG_HEADERS = [
  'RecordKey', 'LineItem', 'Status', 'Created', 'Updated',
  'Run ID', 'Line #', 'Shift', 'Customer Name',
  'Mold ID', 'Mold Description', 'Product Type', 'Resin Lot', 'Item No', 'Item Description',
  'Cavity', 'Test Name', 'DateOfMfg', 'TestDate', 'TestedBy',
  'FreezerTemp', 'DropHeight', 'DropAngle', 'Result', 'FailureDescription', 'Notes', 'Month', 'Year',
];

const DROP_ANGLE_OPTIONS = ['Flat- Bottom', '45 deg- Bottom', 'Flat- Top', '45 deg- Top', 'Side Drop'];

// Literal existing header row of the QC Database (In-Process) log — preserved as-is.
const INPROCESS_LOG_HEADERS = [
  'QC Record #', 'Timestamp Saved', 'Inspection ID', 'Inspection Date', 'Inspection Time',
  'Inspected By', 'Shift', 'Shift Foreman', 'Line #', 'Product Type', 'Mold', 'Color',
  'LOT of Resin', 'Pallet Sequence', 'Sample Date', 'Sample Time', 'Cavity ID', 'Test Type',
  'Measure Index', 'Characteristic Name', 'Unit', 'LSL', 'USL', 'Actual Value', 'Status',
  'Status Detail', 'Visual Notes', 'Source', 'Month', 'Year', 'BatchID', 'Item No.', 'Item Description',
];

// ================= DATA SPREADSHEET ACCESS =================
function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  throw new Error('DATA_SPREADSHEET_ID script property is not set. Run oneTimeSetup() from the script editor first.');
}

function getSettingsSheet_() {
  const sheet = getDb_().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('Settings sheet not found.');
  return sheet;
}

// ================= SETTINGS — dynamic label lookup =================
// The redesigned Settings sheet is read by label text, not fixed cells/columns —
// so it can be reorganized again later without breaking these.
function findSettingsLabel_(label) {
  const st = getSettingsSheet_();
  const lastRow = st.getLastRow(), lastCol = st.getLastColumn();
  const values = st.getRange(1, 1, lastRow, lastCol).getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === label) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function settingsValueRightOf_(label) {
  const pos = findSettingsLabel_(label);
  if (!pos) return '';
  return String(getSettingsSheet_().getRange(pos.row, pos.col + 1).getValue() || '').trim();
}

function settingsColumnBelow_(label, maxScan) {
  const pos = findSettingsLabel_(label);
  if (!pos) return [];
  const st = getSettingsSheet_();
  const out = [];
  let blanks = 0;
  for (let i = 1; i <= (maxScan || 30) && blanks < 6; i++) {
    const v = st.getRange(pos.row + i, pos.col).getValue();
    const s = String(v || '').trim();
    if (!s) { blanks++; continue; }
    blanks = 0;
    out.push(s);
  }
  return out;
}

function settingsTableBelow_(label, numCols, maxScan) {
  const pos = findSettingsLabel_(label);
  if (!pos) return [];
  const st = getSettingsSheet_();
  const lastRow = Math.min(pos.row + (maxScan || 1000), st.getLastRow());
  if (lastRow <= pos.row) return [];
  const data = st.getRange(pos.row + 1, pos.col, lastRow - pos.row, numCols).getValues();
  const out = [];
  let blanks = 0;
  for (const row of data) {
    if (blanks >= 8) break;
    const cells = row.map(v => String(v || '').trim());
    if (cells.every(c => !c)) { blanks++; continue; }
    blanks = 0;
    out.push(cells);
  }
  return out;
}

// ================= MASTER REGISTER LINK (consolidated) =================
function getMasterRegisterId_() {
  const fromSheet = settingsValueRightOf_('Master Register ID (Spec Register):') ||
    settingsValueRightOf_('Master Register ID:');
  return fromSheet || SPEC_REGISTER_ID;
}

function setMasterRegisterId_(idOrUrl) {
  let input = String(idOrUrl || '').trim();
  const m = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) input = m[1];
  SpreadsheetApp.openById(input); // throws if inaccessible
  const pos = findSettingsLabel_('Master Register ID (Spec Register):') || findSettingsLabel_('Master Register ID:');
  if (pos) getSettingsSheet_().getRange(pos.row, pos.col + 1).setValue(input);
  return input;
}

function getNotificationEmails_() {
  const raw = settingsValueRightOf_('Notification Emails (comma-separated):');
  return raw.split(',').map(e => e.trim()).filter(e => e.indexOf('@') > 0);
}

function getInspectorList_()   { return settingsColumnBelow_('QC Technician Name', 30); }
function getForemanList_()     { return settingsColumnBelow_('Shift Foreman', 30); }
function getShiftList_()       { return settingsColumnBelow_('Shift', 30); }
function getPassFailNAList_()  { return settingsColumnBelow_('Pass / Fail / N/A', 30); }
function getStartUpTechList_() { return settingsColumnBelow_('Start-Up Technician Name', 30); }
function getDeviationAuthList_() { return settingsColumnBelow_('Deviation Authorization List', 30); }

/**
 * Reads the Start-Up Verification checklist definitions: [{item, valueType, unit, notes, category}], in sheet order.
 * "Category" is an optional column (e.g. Set-Up, Line Clearance, Artwork & KC#s, Material / Silos) used to group
 * checklist rows on the form — it can live anywhere in the header row, and if it's missing entirely every item
 * comes back with category: '' so the form falls back to one flat, ungrouped table.
 */
function getStartUpItemsList_() {
  const sheet = getDb_().getSheetByName(SU_ITEMS_SHEET_NAME);
  if (!sheet) throw new Error('"' + SU_ITEMS_SHEET_NAME + '" sheet not found.');
  const pos = findHeaderRowAndCol_(sheet, 'Verification Item', 3);
  if (!pos) throw new Error('"Verification Item" header not found in ' + SU_ITEMS_SHEET_NAME + '.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = sheet.getRange(pos.row + 1, pos.col, lastRow - pos.row, 4).getValues();
  const catPos = findHeaderRowAndCol_(sheet, 'Category', 3);
  const categories = catPos ? sheet.getRange(pos.row + 1, catPos.col, lastRow - pos.row, 1).getValues() : null;
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const item = String(row[0] || '').trim();
    if (!item) continue;
    out.push({
      item: item, valueType: String(row[1] || '').trim(), unit: String(row[2] || '').trim(), notes: String(row[3] || '').trim(),
      category: categories ? String(categories[i][0] || '').trim() : '',
    });
  }
  return out;
}

/** Item No. / Item Description list, for traceability fields on both forms. Its own tab. */
function getItemList_() {
  const sheet = getDb_().getSheetByName(ALL_ITEMS_LIST_SHEET_NAME);
  if (!sheet) return [];
  const pos = findHeaderRowAndCol_(sheet, 'Item No.', 3);
  if (!pos) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = sheet.getRange(pos.row + 1, pos.col, lastRow - pos.row, 2).getValues();
  const out = [];
  for (const row of data) {
    const itemNo = String(row[0] || '').trim();
    const description = String(row[1] || '').trim();
    if (!itemNo && !description) continue; // skip blank/section-divider rows (e.g. "Printing")
    if (!itemNo) continue; // no Item No. to key on
    out.push({ itemNo: itemNo, description: description });
  }
  return out;
}
function getItemList() { return getItemList_(); }

// ================= SPEC REGISTER READS (shared by both Drop Freeze + In-Process) =================
// Finds a header cell by exact text within the first few rows of a sheet — robust to
// column reordering, since every Spec Register lookup below reads by header name.
function findHeaderRowAndCol_(sheet, label, maxRows) {
  const lastCol = sheet.getLastColumn();
  const scanRows = Math.min(maxRows || 6, sheet.getLastRow());
  if (scanRows < 1 || lastCol < 1) return null;
  const values = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === label) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

/** Full deduped mold list — reads the register's own "Unique Mold List" helper columns. */
function getAllMoldsList_() {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tab = mr.getSheetByName(ALL_MOLDS_LIST_SHEET);
  if (!tab) throw new Error('"' + ALL_MOLDS_LIST_SHEET + '" tab not found in Spec Register.');
  const pos = findHeaderRowAndCol_(tab, 'Unique Mold List', 5);
  if (!pos) throw new Error('"Unique Mold List" helper column not found in "' + ALL_MOLDS_LIST_SHEET + '".');
  const lastRow = tab.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = tab.getRange(pos.row + 1, pos.col, lastRow - pos.row, 3).getValues();
  const out = [];
  for (const row of data) {
    const moldId = String(row[0] || '').trim();
    if (!moldId) continue;
    out.push({ moldId: moldId, description: String(row[1] || '').trim(), productType: String(row[2] || '').trim() });
  }
  return out;
}
function getAllMoldsList() { return getAllMoldsList_(); }

function getCavityIds_(mold) {
  try {
    const mr = SpreadsheetApp.openById(getMasterRegisterId_());
    const tab = mr.getSheetByName(ALL_MOLDS_LIST_SHEET);
    if (!tab) return [];
    const idPos = findHeaderRowAndCol_(tab, 'Mold ID', 5);
    const cavPos = findHeaderRowAndCol_(tab, 'Cavity IDs', 5);
    if (!idPos || !cavPos) return [];
    const lastRow = tab.getLastRow();
    if (lastRow <= idPos.row) return [];
    const ids = tab.getRange(idPos.row + 1, idPos.col, lastRow - idPos.row, 1).getValues();
    const cavs = tab.getRange(cavPos.row + 1, cavPos.col, lastRow - cavPos.row, 1).getValues();
    const target = String(mold).trim();
    const out = [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === target) {
        const c = String(cavs[i][0] || '').trim();
        if (c) out.push(c);
      }
    }
    return out;
  } catch (e) { return []; }
}

function getSpecsFromMaster_(mold) {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tab = mr.getSheetByName(SPEC_MATRIX_SHEET);
  if (!tab) throw new Error('"' + SPEC_MATRIX_SHEET + '" tab not found in Spec Register.');
  const pos = findHeaderRowAndCol_(tab, 'Mold ID', 6);
  if (!pos) throw new Error('"Mold ID" header not found in ' + SPEC_MATRIX_SHEET + '.');
  const headerRow = pos.row;
  const lastRow = tab.getLastRow();
  if (lastRow <= headerRow) return [];
  const lastCol = tab.getLastColumn();
  const headers = tab.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const idx = function (name) { return headers.indexOf(name); };
  const moldCol = idx('Mold ID'), charCol = idx('Characteristic'), unitCol = idx('Unit'),
    lslCol = idx('LSL'), nomCol = idx('Nominal'), uslCol = idx('USL'), miCol = idx('Measure Index'),
    rejCol = idx('Reject Limit');
  const data = tab.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  const target = String(mold).trim();
  const results = [];
  for (const row of data) {
    if (String(row[moldCol] || '').trim() !== target) continue;
    const lsl = row[lslCol], usl = row[uslCol], nom = row[nomCol];
    const rej = rejCol >= 0 ? row[rejCol] : '';
    const rejPct = (rej === '' || rej === null || rej === undefined) ? null : parseFloat(String(rej).replace('%', ''));
    results.push({
      characteristic: String(row[charCol] || '').trim(),
      unit: String(row[unitCol] || '').trim(),
      measureIndex: row[miCol] !== '' && row[miCol] !== null ? row[miCol] : '',
      lsl: (lsl === '' || lsl === null) ? null : parseFloat(lsl),
      nominal: (nom === '' || nom === null) ? null : parseFloat(nom),
      usl: (usl === '' || usl === null) ? null : parseFloat(usl),
      // Blank Reject Limit = hard LSL/USL only (current behavior). A value (e.g. 1 = 1%) widens
      // a soft band beyond LSL/USL where an excursion is "Needs Review", not an automatic Fail.
      rejectLimitPct: (rejPct === null || isNaN(rejPct)) ? null : rejPct,
    });
  }
  return results;
}

/** Doc No./Revision citation line for spec-driven form headers, read live off the register
 *  rather than hardcoded — stays correct if the revision ever bumps.
 *  Cached (6h) and scoped to just the Register tab — this must stay cheap since it's called
 *  on every page load; a version that scanned every tab in the register once caused the whole
 *  In-Process form to intermittently fail to load (the citation fetch was blocking the same
 *  response as the dropdowns/mold list, and pushed an already-slow cross-spreadsheet open over
 *  the edge). */
function getSpecRegisterCitation_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('specRegisterCitation');
  if (cached) return cached;

  let citation = 'FRM-006-001 Plastics Specification Register';
  try {
    const mr = SpreadsheetApp.openById(getMasterRegisterId_());
    const tab = mr.getSheetByName('Register');
    const scanRows = tab ? Math.min(6, tab.getLastRow()) : 0;
    if (scanRows >= 1) {
      const values = tab.getRange(1, 1, scanRows, tab.getLastColumn()).getValues();
      outer:
      for (const row of values) {
        for (const cell of row) {
          const s = String(cell || '').trim();
          if (/FRM-\d{3}-\d{3}/.test(s) && /revision/i.test(s)) { citation = s; break outer; }
        }
      }
    }
  } catch (e) { /* keep fallback */ }

  cache.put('specRegisterCitation', citation, 21600);
  return citation;
}
function getSpecCitation() { return getSpecRegisterCitation_(); }

/** Reads every Color Specs row for a mold. Each row may carry an Item No. override. */
function readColorSpecRows_(mold) {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tab = mr.getSheetByName(COLOR_SPECS_SHEET);
  if (!tab) return [];
  const pos = findHeaderRowAndCol_(tab, 'Mold ID', 6);
  if (!pos) return [];
  const headerRow = pos.row;
  const lastRow = tab.getLastRow();
  if (lastRow <= headerRow) return [];
  const lastCol = tab.getLastColumn();
  const headers = tab.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const idx = function (name) { return headers.indexOf(name); };
  const moldCol = idx('Mold ID'), itemCol = idx('Item No.'), typeCol = idx('Color Type'), colorCol = idx('Color'),
    lCol = idx('L*'), aCol = idx('a*'), bCol = idx('b*'), deCol = idx('ΔE* Max');
  const data = tab.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  const target = String(mold).trim();
  const out = [];
  for (const row of data) {
    if (String(row[moldCol] || '').trim() !== target) continue;
    out.push({
      itemNo: String(row[itemCol] || '').trim(),
      typeNumber: String(row[typeCol] || '').trim(),
      color: String(row[colorCol] || '').trim(),
      L: row[lCol], a: row[aCol], b: row[bCol],
      deltaEMax: String(row[deCol] || '').trim(),
    });
  }
  return out;
}

/**
 * Color spec for a mold+color, honoring Item No. overrides: an exact Item No.
 * match wins; otherwise falls back to the row with a blank Item No. (the default).
 */
function getColorSpec_(mold, color, itemNo) {
  const colorLower = String(color || '').toLowerCase().trim();
  const itemNoTrim = String(itemNo || '').trim();
  let fallback = null;
  for (const r of readColorSpecRows_(mold)) {
    if (r.color.toLowerCase() !== colorLower) continue;
    if (itemNoTrim && r.itemNo === itemNoTrim) return r;
    if (!r.itemNo && !fallback) fallback = r;
  }
  return fallback;
}

/** Functional Tests rows for a mold, optionally filtered to test names containing `nameFilter`
 *  (case-insensitive substring — e.g. 'Drop Freeze' to exclude the register's other fit tests
 *  like Gauge Fit/Cover Fit/Handle Fit, which aren't in scope for the Drop Freeze module). */
function getFunctionalTestsForMold_(mold, nameFilter) {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tab = mr.getSheetByName(FUNCTIONAL_TESTS_SHEET);
  if (!tab) return [];
  const pos = findHeaderRowAndCol_(tab, 'Mold ID', 6);
  if (!pos) return [];
  const headerRow = pos.row;
  const lastRow = tab.getLastRow();
  if (lastRow <= headerRow) return [];
  const lastCol = tab.getLastColumn();
  const headers = tab.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idx = name => headers.indexOf(name);
  const moldCol = idx('Mold ID'), nameCol = idx('Test Name'), methodCol = idx('Test Method / Description'),
    criteriaCol = idx('Acceptance Criteria'), sampleCol = idx('Sample Size'), freqCol = idx('Frequency'),
    equipCol = idx('Equipment'), resultTypeCol = idx('Result Type');
  const data = tab.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  const target = String(mold).trim();
  const filterLower = String(nameFilter || '').toLowerCase();
  const out = [];
  for (const row of data) {
    if (String(row[moldCol] || '').trim() !== target) continue;
    const testName = String(row[nameCol] || '').trim();
    if (filterLower && testName.toLowerCase().indexOf(filterLower) < 0) continue;
    out.push({
      testName: testName, methodDescription: String(row[methodCol] || '').trim(),
      acceptanceCriteria: String(row[criteriaCol] || '').trim(), sampleSize: String(row[sampleCol] || '').trim(),
      frequency: String(row[freqCol] || '').trim(), equipment: String(row[equipCol] || '').trim(),
      resultType: String(row[resultTypeCol] || '').trim(),
    });
  }
  return out;
}

function getColorOptionsForMold_(mold) {
  const colors = [];
  for (const r of readColorSpecRows_(mold)) {
    if (r.color && colors.indexOf(r.color) < 0) colors.push(r.color);
  }
  return colors;
}

// Dates go over google.script.run as plain strings, not Date objects — the client never needs
// them as real Dates, and many real Date objects in one response is a known flaky spot for
// Apps Script's client bridge (it silently delivers `null` instead of throwing — see Runs work,
// 2026-08-06).
function dateToStr_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

// ================= RUNS (replaces the old Batches concept) =================
// A Run is created once via the Add Run form and stays selectable in In-Process/Drop Freeze
// for as long as it's Active — across shifts, across days — until explicitly Stopped. It also
// carries Qualified/Qualified Timestamp, set by a passing Start-Up Verification submission.
const RUNS_SHEET_NAME = 'Runs';

function getRunsSheet_() {
  const sheet = getDb_().getSheetByName(RUNS_SHEET_NAME);
  if (!sheet) throw new Error('"' + RUNS_SHEET_NAME + '" sheet not found.');
  return sheet;
}

function makeRunId_() {
  const sheet = getRunsSheet_();
  const tz = getDb_().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      const m = String(id || '').match(/^RUN-(\d{8})-(\d{3})$/);
      if (m && m[1] === dateStr) { const seq = Number(m[2]); if (seq > maxSeq) maxSeq = seq; }
    });
  }
  return 'RUN-' + dateStr + '-' + String(maxSeq + 1).padStart(3, '0');
}

function runRowToObject_(row) {
  return {
    runId: row['Run ID'] || '', createdAt: dateToStr_(row['Created At']), shift: row['Shift'] || '',
    status: row['Status'] || '', line: row['Line #'] || '', productType: row['Product Type'] || '',
    resinLot: row['Resin Lot'] || '', moldId: row['Mold ID'] || '', moldDescription: row['Mold Description'] || '',
    item: row['Item'] || '', itemDescription: row['Item Description'] || '', customerName: row['Customer Name'] || '',
    runQty: row['Run Qty'] || '', createdBy: row['Created By'] || '', qualified: row['Qualified'] || '',
    qualifiedTimestamp: dateToStr_(row['Qualified Timestamp']), stoppedAt: dateToStr_(row['Stopped At']),
    lastConfirmed: dateToStr_(row['Last Confirmed']), lastConfirmedBy: row['Last Confirmed By'] || '',
  };
}

/**
 * Creates a new Run. fields: {shift, line, productType, resinLot, moldId, moldDescription,
 * item, itemDescription, customerName, runQty, createdBy}. Returns the created Run.
 */
function createRun_(fields) {
  const sheet = getRunsSheet_();
  const runId = makeRunId_();
  appendObjectsAsRows_(sheet, [{
    'Run ID': runId, 'Created At': dateToStr_(new Date()), 'Shift': fields.shift || '', 'Status': 'Active',
    'Line #': fields.line || '', 'Product Type': fields.productType || '', 'Resin Lot': fields.resinLot || '',
    'Mold ID': fields.moldId || '', 'Mold Description': fields.moldDescription || '',
    'Item': fields.item || '', 'Item Description': fields.itemDescription || '',
    'Customer Name': fields.customerName || '', 'Run Qty': fields.runQty || '', 'Created By': fields.createdBy || '',
  }]);
  return getRun_(runId);
}

function getRun_(runId) {
  const rows = readSheetObjects_(getRunsSheet_());
  const match = rows.find(r => String(r['Run ID'] || '').trim() === String(runId).trim());
  return match ? runRowToObject_(match) : null;
}

/** Active runs, most recently created first — the pool selectable from In-Process/Drop Freeze. */
function getActiveRuns_() {
  const rows = readSheetObjects_(getRunsSheet_());
  return rows.filter(r => String(r['Status'] || '').trim().toLowerCase() === 'active').map(runRowToObject_).reverse();
}

function stopRun_(runId) {
  updateRowWhere_(getRunsSheet_(), 'Run ID', runId, { 'Status': 'Stopped', 'Stopped At': dateToStr_(new Date()) });
  return getRun_(runId);
}

/** Stamps Last Confirmed/Last Confirmed By on every currently-Active run. Returns the count touched. */
function confirmTodaysRuns_(confirmedBy) {
  const sheet = getRunsSheet_();
  const rows = readSheetObjects_(sheet);
  const now = dateToStr_(new Date());
  let count = 0;
  rows.forEach(r => {
    if (String(r['Status'] || '').trim().toLowerCase() === 'active') {
      updateRowWhere_(sheet, 'Run ID', r['Run ID'], { 'Last Confirmed': now, 'Last Confirmed By': confirmedBy || '' });
      count++;
    }
  });
  return count;
}

/** Called when a Start-Up Verification submission passes — marks the Run qualified. */
function qualifyRun_(runId) {
  updateRowWhere_(getRunsSheet_(), 'Run ID', runId, { 'Qualified': 'Yes', 'Qualified Timestamp': dateToStr_(new Date()) });
}

// ================= PUBLIC CLIENT-FACING WRAPPERS (Runs) =================
function createRun(fields) { return createRun_(fields); }
function getActiveRuns() { return getActiveRuns_(); }
function stopRun(runId) { return stopRun_(runId); }
function confirmTodaysRuns(confirmedBy) { return confirmTodaysRuns_(confirmedBy); }

// ================= GENERIC HEADER-KEYED SHEET I/O =================
// Reads every data row into a plain object keyed by column header text —
// robust to legacy/extra columns (e.g. old dashboard helper columns).
function readSheetObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}

// Appends plain objects as rows, matching keys to the sheet's current header order.
// Missing keys are written blank; unknown headers (e.g. legacy formula columns) are left untouched.
function appendObjectsAsRows_(sheet, objects) {
  if (!objects.length) return;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rows = objects.map(obj => headers.map(h => (h && obj.hasOwnProperty(h)) ? obj[h] : ''));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, lastCol).setValues(rows);
}

function deleteRowsWhere_(sheet, keyHeader, keyValue) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const keyCol = headers.indexOf(keyHeader) + 1;
  if (!keyCol) return;
  const keys = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues().flat();
  const rowsToDelete = [];
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i] || '').trim() === String(keyValue).trim()) rowsToDelete.push(i + 2);
  }
  // delete bottom-up so row numbers stay valid
  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
}

// Finds the row where keyHeader === keyValue and patches only the given columns in place —
// unlike deleteRowsWhere_+append, this preserves everything else already on that row. Needed
// for Runs, which get progressively filled in over their lifecycle (create → qualify → stop)
// rather than written once. Returns false if no matching row was found.
function updateRowWhere_(sheet, keyHeader, keyValue, patch) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return false;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const keyCol = headers.indexOf(keyHeader);
  if (keyCol < 0) return false;
  const keys = sheet.getRange(2, keyCol + 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = keys.findIndex(k => String(k || '').trim() === String(keyValue).trim());
  if (rowIdx < 0) return false;
  const sheetRow = rowIdx + 2;
  Object.keys(patch).forEach(h => {
    const col = headers.indexOf(h);
    if (col >= 0) sheet.getRange(sheetRow, col + 1).setValue(patch[h]);
  });
  return true;
}

// ================= MISC HELPERS =================
function isFailValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const up = s.toUpperCase();
  return up.includes('FAIL') || s.includes('❌');
}
