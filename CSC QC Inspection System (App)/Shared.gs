/*************************************************************
 * CSC QC INSPECTION SYSTEM — SHARED CONFIG & HELPERS
 * Consolidated Settings / Master Register / Runs access
 * used by both Drop Freeze and In-Process modules.
 *************************************************************/

// ================= SHEET NAMES =================
const SETTINGS_SHEET_NAME       = 'Settings';
const DROPFREEZE_LOG_SHEET_NAME = 'Drop Freeze Test Data';
const INPROCESS_LOG_SHEET_NAME  = 'QC Inspection Data- Plastics';
const ALL_ITEMS_LIST_SHEET_NAME = 'All Items List- Plastics'; // its own tab, not part of Settings
const ALL_ITEMS_LIST_SHEET_NAME_METALS = 'All Items List- Metals';
const SU_ITEMS_SHEET_NAME = 'Start-Up Verification Items List- Plastics';
const SU_LOG_SHEET_NAME   = 'Start-Up Verification Log- Plastics';
const SU_ITEMS_SHEET_NAME_METALS = 'Start-Up Verification Items List- Metals';
const SU_LOG_SHEET_NAME_METALS   = 'Start-Up Verification Log- Metals';
const METALS_SU_LOG_HEADERS = [
  'Verification Record #', 'Timestamp saved', 'Run ID', 'Verification Date', 'Verification Time',
  'QC Tech Name', 'Shift', 'Foreman', 'Start-Up Tech', 'Line #', 'Run Qty', 'Customer Name',
  'Size ID', 'Item', 'Item Description', 'Month', 'Year',
  'Verification Item', 'Value Type', 'Unit', 'Actual Value', 'Status', 'Notes',
];
const LINE_CONFIG_SHEET_NAME = 'Line Configuration';

const METALS_ENDS_LOG_SHEET_NAME = 'QC Inspection Data- Metals Ends';
const METALS_ENDS_LOG_HEADERS = [
  'QC Record #', 'Timestamp Saved', 'Inspection ID', 'Inspection Date', 'Inspection Time',
  'Inspected By', 'Shift', 'Shift Foreman', 'Line #', 'Machine ID', 'End Size', 'Customer', 'End Description',
  'Test Type', 'Measure Index', 'Characteristic Name', 'Unit', 'LSL', 'USL', 'Actual Value', 'Status', 'Status Detail',
  'Visual Notes', 'Source', 'Month', 'Year', 'Release Decision', 'Justification',
];

// Spec Register (relinked 2026-08 — Spec Matrix / Color Specs / All Molds List
// replaced the old per-product tabs). This is the default; Settings' own
// "Master Register ID" row can override it without a redeploy.
const SPEC_REGISTER_ID = '1rgm9gAnZviUSLKLF1kjbX0P1X5fPvGvKXDg5SvT1htM';

const ALL_MOLDS_LIST_SHEET = 'All Molds List';
const ALL_CANS_SIZE_LIST_SHEET = 'All Cans Size List'; // Metals register — Size ID | Can Description
const SPEC_MATRIX_SHEET    = 'Spec Matrix';
const COLOR_SPECS_SHEET    = 'Color Specs';
const FUNCTIONAL_TESTS_SHEET = 'Functional Tests';

// Header row for the Drop Freeze Test Data log. Run-driven (2026-08-07): shares its
// context columns with the QC Inspection Data / Start-Up Log pattern.
const DROPFREEZE_LOG_HEADERS = [
  'RecordKey', 'LineItem', 'Status', 'Created', 'Updated',
  'Run ID', 'Line #', 'Shift', 'Customer Name',
  'Mold ID', 'Mold Description', 'Product Type', 'Resin Lot', 'Item No', 'Item Description',
  'Cavity', 'Test Name', 'DateOfMfg', 'TestDate', 'TestedBy',
  'FreezerTemp', 'DropHeight', 'DropAngle', 'Result', 'FailureDescription', 'Notes', 'Month', 'Year',
];

const DROP_ANGLE_OPTIONS = ['Flat- Bottom', '45 deg- Bottom', 'Flat- Top', '45 deg- Top', 'Side Drop'];

// In-Process Inspection: choices for resolving a deviation (a measurement outside spec) before
// a record can be submitted — required alongside a written Justification.
const RELEASE_DECISION_OPTIONS = ['Release as is', 'Release with minor deviation', 'Hold', 'Reject'];

// Literal existing header row of the QC Inspection Data- Plastics (In-Process) log — preserved as-is.
// Release Decision/Justification were appended (2026-08) so deviation sign-off gets its own
// columns instead of riding along as fake "Deviation" Characteristic Name rows.
const INPROCESS_LOG_HEADERS = [
  'QC Record #', 'Timestamp Saved', 'Inspection ID', 'Inspection Date', 'Inspection Time',
  'Inspected By', 'Shift', 'Shift Foreman', 'Line #', 'Product Type', 'Mold', 'Color',
  'LOT of Resin', 'Pallet Sequence', 'Sample Date', 'Sample Time', 'Cavity ID', 'Test Type',
  'Measure Index', 'Characteristic Name', 'Unit', 'LSL', 'USL', 'Actual Value', 'Status',
  'Status Detail', 'Visual Notes', 'Source', 'Month', 'Year', 'BatchID', 'Item No.', 'Item Description',
  'Release Decision', 'Justification',
];

// ================= DATA SPREADSHEET ACCESS =================
// Memoized per execution — SpreadsheetApp.openById() is one of the slowest calls available,
// and every helper in this file used to call getDb_() independently, re-opening the same
// spreadsheet many times over during a single form load. Globals reset on every fresh
// execution, so this can't leak stale data across requests.
let _dbCache_ = null;
function getDb_() {
  if (_dbCache_) return _dbCache_;
  const id = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
  if (id) { _dbCache_ = SpreadsheetApp.openById(id); return _dbCache_; }
  const active = SpreadsheetApp.getActive();
  if (active) { _dbCache_ = active; return _dbCache_; }
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
// fromRow lets a caller skip past an earlier section that reuses the same label text
// (e.g. "QC Technician Name" appears once under Plastics' dropdown lists, again under Metals').

// Memoized per execution, like _dbCache_ above. Every dropdown/label lookup on a form
// (QC Tech, Foreman, Shift, Start-Up Tech, etc.) used to call findSettingsLabel_ separately,
// and each call re-fetched the ENTIRE Settings sheet with its own getRange().getValues() —
// on the Metals Start-Up form alone that was ~8 full-sheet reads just to locate labels,
// which was the real source of the multi-second load. Now the sheet is fetched once and every
// lookup for the rest of this execution just scans the in-memory array.
let _settingsGridCache_ = null;
function getSettingsGrid_() {
  if (_settingsGridCache_) return _settingsGridCache_;
  const st = getSettingsSheet_();
  const lastRow = st.getLastRow(), lastCol = st.getLastColumn();
  _settingsGridCache_ = {
    values: (lastRow > 0 && lastCol > 0) ? st.getRange(1, 1, lastRow, lastCol).getValues() : [],
    lastRow: lastRow, lastCol: lastCol,
  };
  return _settingsGridCache_;
}

function findSettingsLabel_(label, fromRow) {
  const grid = getSettingsGrid_();
  const startRow = fromRow || 1;
  if (startRow > grid.lastRow) return null;
  for (let r = startRow - 1; r < grid.values.length; r++) {
    const row = grid.values[r];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c]).trim() === label) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function settingsValueRightOf_(label) {
  const pos = findSettingsLabel_(label);
  if (!pos) return '';
  return String(getSettingsSheet_().getRange(pos.row, pos.col + 1).getValue() || '').trim();
}

/** Row of the "Drop Down Lists- Metals" section header, if Settings still keeps a separate
 *  Metals section. Returns null when it doesn't (e.g. merged into one shared "Drop Down
 *  Lists" section, as of the 2026-08 Settings cleanup) — callers fall back to the same
 *  lookup Plastics uses instead of erroring, since that's exactly what a shared section means.
 *  Must not throw: it's called while building getStartUpVerificationFormData()'s single
 *  response object, and any throw there blanks out every field on the form, not just this one. */
function getMetalsDropdownSectionRow_() {
  const pos = findSettingsLabel_('Drop Down Lists- Metals');
  return pos ? pos.row : null;
}

function settingsColumnBelow_(label, maxScan, fromRow) {
  const pos = findSettingsLabel_(label, fromRow);
  if (!pos) return [];
  // Reads off the same cached grid findSettingsLabel_ just used instead of issuing its own
  // getRange() round trip — see getSettingsGrid_ above.
  const grid = getSettingsGrid_();
  const scan = Math.min(maxScan || 30, Math.max(grid.lastRow - pos.row, 0));
  if (scan <= 0) return [];
  const out = [];
  let blanks = 0;
  for (let i = 0; i < scan && blanks < 6; i++) {
    const row = grid.values[pos.row + i] || [];
    const s = String(row[pos.col - 1] || '').trim();
    if (!s) { blanks++; continue; }
    blanks = 0;
    out.push(s);
  }
  return out;
}

function settingsTableBelow_(label, numCols, maxScan) {
  const pos = findSettingsLabel_(label);
  if (!pos) return [];
  const grid = getSettingsGrid_();
  const lastRow = Math.min(pos.row + (maxScan || 1000), grid.lastRow);
  if (lastRow <= pos.row) return [];
  const out = [];
  let blanks = 0;
  for (let r = pos.row; r < lastRow; r++) {
    if (blanks >= 8) break;
    const row = grid.values[r] || [];
    const cells = [];
    for (let c = pos.col - 1; c < pos.col - 1 + numCols; c++) cells.push(String(row[c] || '').trim());
    if (cells.every(c => !c)) { blanks++; continue; }
    blanks = 0;
    out.push(cells);
  }
  return out;
}

// ================= MASTER REGISTER LINK (consolidated) =================
// Memoized per execution, per department — resolving the ID re-scans the Settings sheet, and
// SpreadsheetApp.openById() on the (separate, external) Spec Register is one of the slowest
// calls available. A single In-Process cavity/spec/color/functional-test lookup used to reopen
// the register up to 4 times over; now it opens once and every helper below reuses it.
const _registerCache_ = {};
function getMasterRegisterId_(department) {
  const key = department === 'Metals' ? 'Metals' : 'Plastics';
  if (_registerCache_[key] && _registerCache_[key].id) return _registerCache_[key].id;
  let id;
  if (department === 'Metals') {
    id = settingsValueRightOf_('Master Register ID (Spec Register)- Metals:');
    if (!id) throw new Error('Metals Master Register ID not set in Settings yet.');
  } else {
    id = settingsValueRightOf_('Master Register ID (Spec Register):') ||
      settingsValueRightOf_('Master Register ID:') || SPEC_REGISTER_ID;
  }
  _registerCache_[key] = { id: id };
  return id;
}

function getMasterRegister_(department) {
  const key = department === 'Metals' ? 'Metals' : 'Plastics';
  const id = getMasterRegisterId_(department);
  if (_registerCache_[key].ss) return _registerCache_[key].ss;
  _registerCache_[key].ss = SpreadsheetApp.openById(id);
  return _registerCache_[key].ss;
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

function getInspectorList_(department)   { return settingsColumnBelow_('QC Technician Name', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }
function getForemanList_(department)     { return settingsColumnBelow_('Shift Foreman', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }
function getShiftList_(department)       { return settingsColumnBelow_('Shift', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }
function getPassFailNAList_(department)  { return settingsColumnBelow_('Pass / Fail / N/A', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }
function getStartUpTechList_(department) { return settingsColumnBelow_('Start-Up Technician Name', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }
function getDeviationAuthList_(department) { return settingsColumnBelow_('Deviation Authorization List', 30, department === 'Metals' ? getMetalsDropdownSectionRow_() : undefined); }

/**
 * Reads the Start-Up Verification checklist definitions: [{item, valueType, unit, notes, category, number}],
 * in sheet order. "Category" is an optional column (e.g. Set-Up, Line Clearance, Artwork & KC#s, Material /
 * Silos) used to group checklist rows on the form — it can live anywhere in the header row, and if it's
 * missing entirely every item comes back with category: '' so the form falls back to one flat, ungrouped
 * table. "No." is likewise an optional column holding a manually-assigned display number for the item —
 * missing means number: '' and the form just doesn't show one.
 */
function getStartUpItemsList_(department) {
  const sheetName = department === 'Metals' ? SU_ITEMS_SHEET_NAME_METALS : SU_ITEMS_SHEET_NAME;
  const sheet = getDb_().getSheetByName(sheetName);
  if (!sheet) throw new Error('"' + sheetName + '" sheet not found.');
  const pos = findHeaderRowAndCol_(sheet, 'Verification Item', 3);
  if (!pos) throw new Error('"Verification Item" header not found in ' + sheetName + '.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = sheet.getRange(pos.row + 1, pos.col, lastRow - pos.row, 4).getValues();
  const catPos = findHeaderRowAndCol_(sheet, 'Category', 3);
  const categories = catPos ? sheet.getRange(pos.row + 1, catPos.col, lastRow - pos.row, 1).getValues() : null;
  const numPos = findHeaderRowAndCol_(sheet, 'No.', 3);
  const numbers = numPos ? sheet.getRange(pos.row + 1, numPos.col, lastRow - pos.row, 1).getValues() : null;
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const item = String(row[0] || '').trim();
    if (!item) continue;
    out.push({
      item: item, valueType: String(row[1] || '').trim(), unit: String(row[2] || '').trim(), notes: String(row[3] || '').trim(),
      category: categories ? String(categories[i][0] || '').trim() : '',
      number: numbers ? String(numbers[i][0] || '').trim() : '',
    });
  }
  // "No." is the real sort key, not just a display label — sorted here (once, department-agnostic)
  // rather than client-side so category order and the form's Yes/No → Number → Other sub-grouping
  // both inherit it for free. Numbered items sort ascending; unnumbered items keep their original
  // sheet-row order and always sort after every numbered item (never jump ahead just for being
  // blank). Stable tie-break on original index so equal/blank numbers don't get reshuffled.
  const withIndex = out.map((it, i) => ({ it: it, i: i, num: it.number === '' ? null : parseFloat(it.number) }));
  withIndex.sort((a, b) => {
    if (a.num === null && b.num === null) return a.i - b.i;
    if (a.num === null) return 1;
    if (b.num === null) return -1;
    if (a.num !== b.num) return a.num - b.num;
    return a.i - b.i;
  });
  return withIndex.map(x => x.it);
}

/**
 * Distinct Line # values for a department, from the Line Configuration sheet
 * (Department | Line # | Equipment Code | Equipment Description — one row per piece of
 * equipment, so a line with several equipment rows still only yields one Line # here).
 * Admin-managed by hand for now; sorted numerically when the Line # is a plain number.
 */
function getLinesForDepartment_(department) {
  const sheet = getDb_().getSheetByName(LINE_CONFIG_SHEET_NAME);
  if (!sheet) return [];
  const seen = {};
  const out = [];
  readSheetObjects_(sheet).forEach(row => {
    const dept = String(row['Department'] || '').trim();
    const line = String(row['Line #'] || '').trim();
    if (!line || dept !== department || seen[line]) return;
    seen[line] = true;
    out.push(line);
  });
  out.sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b);
  });
  return out;
}

/** Item No. / Item Description list, for traceability fields on both forms. Its own tab per department. */
function getItemList_(department) {
  const sheetName = department === 'Metals' ? ALL_ITEMS_LIST_SHEET_NAME_METALS : ALL_ITEMS_LIST_SHEET_NAME;
  const sheet = getDb_().getSheetByName(sheetName);
  if (!sheet) return [];
  const pos = findHeaderRowAndCol_(sheet, 'Item No.', 3) || findHeaderRowAndCol_(sheet, 'Item No', 3);
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

/**
 * Full deduped mold list — reads the register's "Unique Mold List" helper column plus
 * Product Type right next to it. No description column anymore (removed from the register) —
 * every caller now shows Mold ID / Product Type only.
 */
function getAllMoldsList_() {
  const mr = getMasterRegister_();
  const tab = mr.getSheetByName(ALL_MOLDS_LIST_SHEET);
  if (!tab) throw new Error('"' + ALL_MOLDS_LIST_SHEET + '" tab not found in Spec Register.');
  const pos = findHeaderRowAndCol_(tab, 'Unique Mold List', 5);
  if (!pos) throw new Error('"Unique Mold List" helper column not found in "' + ALL_MOLDS_LIST_SHEET + '".');
  const lastRow = tab.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = tab.getRange(pos.row + 1, pos.col, lastRow - pos.row, 2).getValues();
  const out = [];
  for (const row of data) {
    const moldId = String(row[0] || '').trim();
    if (!moldId) continue;
    out.push({ moldId: moldId, productType: String(row[1] || '').trim() });
  }
  return out;
}
function getAllMoldsList() { return getAllMoldsList_(); }

/**
 * Metals equivalent of getAllMoldsList_ — reads the deduped "Unique Can Size List" helper
 * column (Size ID repeats in the raw list due to customer-specific Spec Matrix rows) plus
 * Product Type right next to it. No description column here — too specific/near-duplicate
 * of the Item Description already captured separately on Add Run.
 */
function getAllSizeCansList_() {
  const mr = getMasterRegister_('Metals');
  const tab = mr.getSheetByName(ALL_CANS_SIZE_LIST_SHEET);
  if (!tab) throw new Error('"' + ALL_CANS_SIZE_LIST_SHEET + '" tab not found in Metals Spec Register.');
  const pos = findHeaderRowAndCol_(tab, 'Unique Can Size List', 5);
  if (!pos) throw new Error('"Unique Can Size List" helper column not found in "' + ALL_CANS_SIZE_LIST_SHEET + '".');
  const lastRow = tab.getLastRow();
  if (lastRow <= pos.row) return [];
  const data = tab.getRange(pos.row + 1, pos.col, lastRow - pos.row, 2).getValues();
  const out = [];
  for (const row of data) {
    const sizeId = String(row[0] || '').trim();
    if (!sizeId) continue;
    out.push({ sizeId: sizeId, productType: String(row[1] || '').trim() });
  }
  return out;
}
function getAllSizeCansList() { return getAllSizeCansList_(); }

function getCavityIds_(mold) {
  try {
    const mr = getMasterRegister_();
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
  const mr = getMasterRegister_();
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

/**
 * Metals Ends spec register reads — the Metals Master Register's own "Spec Matrix" tab covers
 * both End and (eventually) Body components in one sheet, distinguished by a "Product Type"
 * column ('End' / 'Body'), unlike Plastics' single-purpose Spec Matrix keyed by Mold ID alone.
 * No dedicated "All Ends Size List" helper tab exists (unlike All Molds List / All Cans Size
 * List), so the End Size list is derived here by deduping Spec Matrix rows directly.
 */
function getEndsSpecMatrixTab_() {
  const mr = getMasterRegister_('Metals');
  const tab = mr.getSheetByName(SPEC_MATRIX_SHEET);
  if (!tab) throw new Error('"' + SPEC_MATRIX_SHEET + '" tab not found in Metals Spec Register.');
  return tab;
}

/** [{sizeId, customer, description}] for every distinct End Size on file, in first-seen order. */
function getEndSizeList_() {
  const tab = getEndsSpecMatrixTab_();
  const pos = findHeaderRowAndCol_(tab, 'Size ID', 6);
  if (!pos) return [];
  const lastRow = tab.getLastRow(), lastCol = tab.getLastColumn();
  if (lastRow <= pos.row) return [];
  const headers = tab.getRange(pos.row, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idx = name => headers.indexOf(name);
  const sizeCol = idx('Size ID'), custCol = idx('Customer'), descCol = idx('Description'), prodCol = idx('Product Type');
  const data = tab.getRange(pos.row + 1, 1, lastRow - pos.row, lastCol).getValues();
  const seen = {};
  const out = [];
  data.forEach(row => {
    if (String(row[prodCol] || '').trim() !== 'End') return;
    const sizeId = String(row[sizeCol] || '').trim();
    if (!sizeId || seen[sizeId]) return;
    seen[sizeId] = true;
    out.push({ sizeId: sizeId, customer: String(row[custCol] || '').trim(), description: String(row[descCol] || '').trim() });
  });
  return out;
}
function getEndSizeList() { return getEndSizeList_(); }

/** Dimensional specs for one End Size — same shape as getSpecsFromMaster_'s results, filtered
 *  to this Size ID + Product Type='End'. */
function getSpecsForEndSize_(sizeId) {
  const tab = getEndsSpecMatrixTab_();
  const pos = findHeaderRowAndCol_(tab, 'Size ID', 6);
  if (!pos) throw new Error('"Size ID" header not found in Metals ' + SPEC_MATRIX_SHEET + '.');
  const lastRow = tab.getLastRow(), lastCol = tab.getLastColumn();
  if (lastRow <= pos.row) return [];
  const headers = tab.getRange(pos.row, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idx = name => headers.indexOf(name);
  const sizeCol = idx('Size ID'), prodCol = idx('Product Type'), charCol = idx('Characteristic'), unitCol = idx('Unit'),
    lslCol = idx('LSL'), nomCol = idx('Nominal'), uslCol = idx('USL'), miCol = idx('Measure Index'), rejCol = idx('Reject Limit');
  const data = tab.getRange(pos.row + 1, 1, lastRow - pos.row, lastCol).getValues();
  const target = String(sizeId).trim();
  const results = [];
  data.forEach(row => {
    if (String(row[sizeCol] || '').trim() !== target) return;
    if (String(row[prodCol] || '').trim() !== 'End') return;
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
      rejectLimitPct: (rejPct === null || isNaN(rejPct)) ? null : rejPct,
    });
  });
  return results;
}

/** Equipment Code + Description options for one Line, from Line Configuration — backs the
 *  Machine ID dropdown on Metals Ends In-Process (per-Line, the way Cavity IDs are per-Mold on
 *  Plastics). Empty until real Ends-press equipment rows exist there — add them by hand the
 *  same way every other Line Configuration row is added. */
function getMachineIdsForLine_(department, line) {
  const sheet = getDb_().getSheetByName(LINE_CONFIG_SHEET_NAME);
  if (!sheet) return [];
  const out = [];
  readSheetObjects_(sheet).forEach(row => {
    if (String(row['Department'] || '').trim() !== department) return;
    if (String(row['Line #'] || '').trim() !== String(line).trim()) return;
    const code = String(row['Equipment Code'] || '').trim();
    if (!code) return;
    out.push({ code: code, description: String(row['Equipment Description'] || '').trim() });
  });
  return out;
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
    const mr = getMasterRegister_();
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
  const mr = getMasterRegister_();
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
  const mr = getMasterRegister_();
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
// 2026-08-06). This only applies to values crossing that bridge to the client, though — audit
// timestamps written straight into a sheet cell (Created At, Stopped At, Timestamp saved, etc.)
// should be plain `new Date()`, not dateToStr_(new Date()): Range.setValues() isn't the client
// bridge, and a real Date cell displays as a normal date/time instead of raw ISO text
// (2026-08-07T14:46:39.472Z) and stays sortable/filterable. dateToStr_ still runs on the way
// back OUT to the client (e.g. runRowToObject_), so the flaky-bridge constraint above still holds.
function dateToStr_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

// ================= RUNS (replaces the old Batches concept) =================
// A Run is created once via the Add Run form and stays selectable in In-Process/Drop Freeze
// for as long as it's Active — across shifts, across days — until explicitly Stopped. It also
// carries Qualified/Qualified Timestamp, set by a passing Start-Up Verification submission.
const RUNS_SHEET_NAME = 'Runs - Plastics';
const RUNS_SHEET_NAME_METALS = 'Runs - Metals';
const METALS_RUNS_HEADERS = [
  'Run ID', 'Created At', 'Shift', 'Status', 'Line #', 'Product Type', 'Material Lot',
  'Size ID', 'Can Description', 'Item', 'Item Description', 'Customer Name', 'Run Qty', 'Created By',
  'Qualified', 'Qualified Timestamp', 'Stopped At', 'Last Confirmed', 'Last Confirmed By',
];

function getRunsSheetName_(department) { return department === 'Metals' ? RUNS_SHEET_NAME_METALS : RUNS_SHEET_NAME; }

function getRunsSheet_(department) {
  const name = getRunsSheetName_(department);
  const sheet = getDb_().getSheetByName(name);
  if (!sheet) throw new Error('"' + name + '" sheet not found.');
  return sheet;
}

function makeRunId_(department) {
  return makeSequentialId_(getRunsSheet_(department), 'Run ID', 'RUN');
}

// Reads both Plastics' and Metals' domain-specific columns — whichever set the row's sheet
// doesn't have simply comes back blank, so this stays a single function for both departments.
function runRowToObject_(row) {
  return {
    runId: row['Run ID'] || '', createdAt: dateToStr_(row['Created At']), shift: row['Shift'] || '',
    status: row['Status'] || '', line: row['Line #'] || '', productType: row['Product Type'] || '',
    resinLot: row['Resin Lot'] || '', moldId: row['Mold ID'] || '', moldDescription: row['Mold Description'] || '',
    materialLot: row['Material Lot'] || '', sizeId: row['Size ID'] || '', canDescription: row['Can Description'] || '',
    item: row['Item'] || '', itemDescription: row['Item Description'] || '', customerName: row['Customer Name'] || '',
    runQty: row['Run Qty'] || '', createdBy: row['Created By'] || '', qualified: row['Qualified'] || '',
    qualifiedTimestamp: dateToStr_(row['Qualified Timestamp']), stoppedAt: dateToStr_(row['Stopped At']),
    lastConfirmed: dateToStr_(row['Last Confirmed']), lastConfirmedBy: row['Last Confirmed By'] || '',
  };
}

/**
 * Creates a new Run. fields: {shift, line, productType, resinLot, moldId, moldDescription,
 * materialLot, sizeId, canDescription, item, itemDescription, customerName, runQty, createdBy}.
 * Returns the created Run. department defaults to Plastics.
 */
function createRun_(fields, department) {
  const sheet = getRunsSheet_(department);
  const runId = makeRunId_(department);
  appendObjectsAsRows_(sheet, [{
    'Run ID': runId, 'Created At': new Date(), 'Shift': fields.shift || '', 'Status': 'Active',
    'Line #': fields.line || '', 'Product Type': fields.productType || '', 'Resin Lot': fields.resinLot || '',
    'Mold ID': fields.moldId || '', 'Mold Description': fields.moldDescription || '',
    'Material Lot': fields.materialLot || '', 'Size ID': fields.sizeId || '', 'Can Description': fields.canDescription || '',
    'Item': fields.item || '', 'Item Description': fields.itemDescription || '',
    'Customer Name': fields.customerName || '', 'Run Qty': fields.runQty || '', 'Created By': fields.createdBy || '',
  }]);
  return getRun_(runId, department);
}

function getRun_(runId, department) {
  const rows = readSheetObjects_(getRunsSheet_(department));
  const match = rows.find(r => String(r['Run ID'] || '').trim() === String(runId).trim());
  return match ? runRowToObject_(match) : null;
}

/** Active runs, most recently created first — the pool selectable from In-Process/Drop Freeze. */
function getActiveRuns_(department) {
  const rows = readSheetObjects_(getRunsSheet_(department));
  return rows.filter(r => String(r['Status'] || '').trim().toLowerCase() === 'active').map(runRowToObject_).reverse();
}

function stopRun_(runId, department) {
  updateRowWhere_(getRunsSheet_(department), 'Run ID', runId, { 'Status': 'Stopped', 'Stopped At': new Date() });
  return getRun_(runId, department);
}

/** Stamps Last Confirmed/Last Confirmed By on every currently-Active run. Returns the count touched. */
function confirmTodaysRuns_(confirmedBy, department) {
  const sheet = getRunsSheet_(department);
  const rows = readSheetObjects_(sheet);
  const now = new Date();
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
function qualifyRun_(runId, department) {
  updateRowWhere_(getRunsSheet_(department), 'Run ID', runId, { 'Qualified': 'Yes', 'Qualified Timestamp': new Date() });
}

// ================= PUBLIC CLIENT-FACING WRAPPERS (Runs) =================
function createRun(fields) { return createRun_(fields); }
function getActiveRuns() { return getActiveRuns_(); }
function stopRun(runId) { return stopRun_(runId); }
function confirmTodaysRuns(confirmedBy) { return confirmTodaysRuns_(confirmedBy); }

// ================= SEQUENTIAL RECORD IDs (Run/QC/SUV/PFA) =================
// Shared by makeRunId_ (here), InProcess.gs's makeRecordID_, DropFreeze.gs's
// makeDailyRecordKey_, and StartUpVerification.gs's makeVerificationRecordId_/makePfaId_ —
// one PREFIX-yyMMdd-N scheme instead of four near-identical copies. The sequence resets each
// calendar year (not daily): scans `column` on `sheet` for existing PREFIX IDs whose embedded
// 2-digit year matches baseDate's, and returns one past the highest sequence number found.
// baseDate defaults to now; Drop Freeze backdates it to the sample's Date of Mfg.
function makeSequentialId_(sheet, column, prefix, baseDate) {
  const tz = getDb_().getSpreadsheetTimeZone();
  const when = (baseDate instanceof Date && !isNaN(baseDate.getTime())) ? baseDate : new Date();
  const dateStr = Utilities.formatDate(when, tz, 'yyMMdd');
  const yy = dateStr.substring(0, 2);
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf(column);
    if (col >= 0) {
      const re = new RegExp('^' + prefix + '-(\\d{2})\\d{4}-(\\d+)$');
      const ids = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues().flat();
      ids.forEach(id => {
        const m = String(id || '').match(re);
        if (m && m[1] === yy) { const seq = Number(m[2]); if (seq > maxSeq) maxSeq = seq; }
      });
    }
  }
  return prefix + '-' + dateStr + '-' + (maxSeq + 1);
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
