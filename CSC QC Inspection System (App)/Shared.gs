/*************************************************************
 * CSC QC INSPECTION SYSTEM — SHARED CONFIG & HELPERS
 * Consolidated Settings / Master Register / Batches access
 * used by both Drop Freeze and In-Process modules.
 *************************************************************/

// ================= SHEET NAMES =================
const SETTINGS_SHEET_NAME       = 'Settings';
const BATCHES_SHEET_NAME        = 'Batches';
const DROPFREEZE_LOG_SHEET_NAME = 'Test Data';
const INPROCESS_LOG_SHEET_NAME  = 'QC Database';
const ALL_ITEMS_LIST_SHEET_NAME = 'All Items List'; // its own tab, not part of Settings

// Spec Register (relinked 2026-08 — Spec Matrix / Color Specs / All Molds List
// replaced the old per-product tabs). This is the default; Settings' own
// "Master Register ID" row can override it without a redeploy.
const SPEC_REGISTER_ID = '1rgm9gAnZviUSLKLF1kjbX0P1X5fPvGvKXDg5SvT1htM';

const ALL_MOLDS_LIST_SHEET = 'All Molds List';
const SPEC_MATRIX_SHEET    = 'Spec Matrix';
const COLOR_SPECS_SHEET    = 'Color Specs';

const BATCHES_HEADERS = [
  'BatchID', 'Date', 'Line', 'Shift', 'ProductType', 'Mold',
  'Pallet', 'ResinLot', 'Inspector', 'CreatedAt', 'Status',
];

// Clean header row for a brand-new Test Data (Drop Freeze) log — this working
// copy has no Drop Freeze data yet, so we're free to define it cleanly.
const DROPFREEZE_LOG_HEADERS = [
  'RecordKey', 'LineItem', 'Status', 'Created', 'Updated',
  'LineNum', 'DateOfMfg', 'Shift', 'Pallet', 'ToolCode', 'ProductType', 'Cavity', 'ResinID', 'TestType',
  'ItemNo', 'ItemDescription',
  'TestDate', 'TestedBy', 'FreezerTemp', 'DropHeight', 'DropAngle',
  'Result', 'FailureDescription', 'Notes', 'BatchID',
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

function getInspectorList_()   { return settingsColumnBelow_('Inspected By', 30); }
function getForemanList_()     { return settingsColumnBelow_('Shift Foreman', 30); }
function getShiftList_()       { return settingsColumnBelow_('Shift', 30); }
function getPassFailNAList_()  { return settingsColumnBelow_('Pass / Fail / N/A', 30); }

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

function getSpecsFromMaster_(productType, mold) {
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
    lslCol = idx('LSL'), nomCol = idx('Nominal'), uslCol = idx('USL'), miCol = idx('Measure Index');
  const data = tab.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  const target = String(mold).trim();
  const results = [];
  for (const row of data) {
    if (String(row[moldCol] || '').trim() !== target) continue;
    const lsl = row[lslCol], usl = row[uslCol], nom = row[nomCol];
    results.push({
      characteristic: String(row[charCol] || '').trim(),
      unit: String(row[unitCol] || '').trim(),
      measureIndex: row[miCol] !== '' && row[miCol] !== null ? row[miCol] : '',
      lsl: (lsl === '' || lsl === null) ? null : parseFloat(lsl),
      nominal: (nom === '' || nom === null) ? null : parseFloat(nom),
      usl: (usl === '' || usl === null) ? null : parseFloat(usl),
    });
  }
  return results;
}

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

function getColorOptionsForMold_(mold) {
  const colors = [];
  for (const r of readColorSpecRows_(mold)) {
    if (r.color && colors.indexOf(r.color) < 0) colors.push(r.color);
  }
  return colors;
}

// ================= BATCHES (shared batch-context, fixes double entry) =================
function getBatchesSheet_() {
  const sheet = getDb_().getSheetByName(BATCHES_SHEET_NAME);
  if (!sheet) throw new Error('Batches sheet not found. Run oneTimeSetup() first.');
  return sheet;
}

function makeBatchId_() {
  const sheet = getBatchesSheet_();
  const tz = getDb_().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      const m = String(id || '').match(/^B-(\d{8})-(\d{3})$/);
      if (m && m[1] === dateStr) {
        const seq = Number(m[2]);
        if (seq > maxSeq) maxSeq = seq;
      }
    });
  }
  return 'B-' + dateStr + '-' + String(maxSeq + 1).padStart(3, '0');
}

/**
 * Creates a new batch record. fields: {line, shift, productType, mold, pallet, resinLot, inspector}
 * Returns the created batch object (including its new BatchID).
 */
function createBatch_(fields) {
  const sheet = getBatchesSheet_();
  const batchId = makeBatchId_();
  const now = new Date();
  const row = [
    batchId, now, fields.line || '', fields.shift || '', fields.productType || '',
    fields.mold || '', fields.pallet || '', fields.resinLot || '', fields.inspector || '',
    now, 'Open',
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return batchToObject_(row);
}

function batchToObject_(row) {
  return {
    batchId: row[0], date: row[1], line: row[2], shift: row[3], productType: row[4],
    mold: row[5], pallet: row[6], resinLot: row[7], inspector: row[8],
    createdAt: row[9], status: row[10],
  };
}

/** Returns open batches, most recent first — used to populate the "load existing batch" picker. */
function getOpenBatches_() {
  const sheet = getBatchesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, BATCHES_HEADERS.length).getValues();
  return rows
    .filter(r => String(r[10]).trim().toLowerCase() === 'open')
    .map(batchToObject_)
    .reverse();
}

function getBatch_(batchId) {
  const sheet = getBatchesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, BATCHES_HEADERS.length).getValues();
  const match = rows.find(r => String(r[0]).trim() === String(batchId).trim());
  return match ? batchToObject_(match) : null;
}

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

// ================= PUBLIC CLIENT-FACING WRAPPERS =================
function createBatch(fields) { return createBatch_(fields); }
function getOpenBatches() { return getOpenBatches_(); }

// ================= MISC HELPERS =================
function isFailValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const up = s.toUpperCase();
  return up.includes('FAIL') || s.includes('❌');
}
