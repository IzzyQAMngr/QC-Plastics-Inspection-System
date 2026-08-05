/*************************************************************
 * FRM-006-001 SPEC REGISTER — auto Mold Description lookup +
 * soft mold-group row shading, for Spec Matrix / Color Specs /
 * Functional Tests. Source of truth for Mold ID + Description
 * is the "All Molds List" tab.
 *************************************************************/

const ALL_MOLDS_SHEET_NAME = 'All Molds List';
const TARGET_SHEETS = ['Spec Matrix', 'Color Specs', 'Functional Tests'];
const HEADER_ROW = 4;
const DATA_START_ROW = 5;

const MOLD_ID_HEADER = 'Mold ID';
const MOLD_DESC_HEADER = 'Mold Description';

const DESC_GREY = '#E0E0E0';
const SHADE_A = '#FFFFFF';
const SHADE_B = '#F4F4F4';
const NOT_FOUND_BG = '#FDECEA';

const PROTECTION_TAG = 'AUTO_MOLD_DESCRIPTION — populated by script, do not edit directly';
const PROTECTED_ROW_COUNT = 2000; // generous headroom so newly-added rows are covered without re-running setup
const EDIT_HANDLER = 'handleMoldIdEdit_';

// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Spec Register Tools')
    .addItem('🔄 Refresh All Descriptions + Shading', 'refreshAll')
    .addSeparator()
    .addItem('⚙️ Run One-Time Setup (protect + trigger + backfill)', 'oneTimeSetup')
    .addToUi();
}

// ================= ONE-TIME SETUP =================
// Run once from the Apps Script editor (or via the menu once authorized).
// Installs an installable onEdit trigger (so it still fires — and can still
// write to the protected column — no matter who edits the sheet), protects
// the Description columns, and backfills everything immediately.
function oneTimeSetup() {
  ensureEditTrigger_();
  protectDescriptionColumns_();
  refreshAll();
  SpreadsheetApp.getActive().toast('Setup complete: trigger installed, Description columns protected, all descriptions populated.');
}

function ensureEditTrigger_() {
  const ss = SpreadsheetApp.getActive();
  const already = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === EDIT_HANDLER && t.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
  if (!already) {
    ScriptApp.newTrigger(EDIT_HANDLER).forSpreadsheet(ss).onEdit().create();
  }
}

function protectDescriptionColumns_() {
  TARGET_SHEETS.forEach(function (name) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (!sheet) return;
    const descCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_DESC_HEADER);
    if (!descCol) return;

    // Remove any protection we previously created here so re-running setup is safe.
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      if (p.getDescription() === PROTECTION_TAG) p.remove();
    });

    const range = sheet.getRange(DATA_START_ROW, descCol, PROTECTED_ROW_COUNT, 1);
    const protection = range.protect();
    protection.setDescription(PROTECTION_TAG);
    protection.setWarningOnly(false);
    try {
      protection.removeEditors(protection.getEditors());
      protection.addEditor(Session.getEffectiveUser());
    } catch (e) { /* some accounts can't edit the editor list — protection still applies */ }
  });
}

// ================= ON EDIT (installable — see ensureEditTrigger_) =================
// NOTE: intentionally NOT named "onEdit" — that would make it an auto-registered
// simple trigger too, which runs with the editing user's own permissions and would
// get blocked by the protected Description column. Installable triggers run with
// the authorizing user's permissions instead, so this can write through protection
// no matter who actually made the edit.
function handleMoldIdEdit_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (TARGET_SHEETS.indexOf(sheet.getName()) === -1) return;

  const moldIdCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_ID_HEADER);
  if (!moldIdCol) return;

  const editStartCol = e.range.getColumn();
  const editEndCol = editStartCol + e.range.getNumColumns() - 1;
  if (editStartCol > moldIdCol || editEndCol < moldIdCol) return; // edit didn't touch Mold ID column

  refreshSheet_(sheet);
}

// ================= MANUAL / ONE-TIME REFRESH =================
function refreshAll() {
  TARGET_SHEETS.forEach(function (name) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) refreshSheet_(sheet);
  });
  SpreadsheetApp.getActive().toast('Descriptions + mold shading refreshed on: ' + TARGET_SHEETS.join(', '));
}

// ================= CORE =================
function refreshSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  shadeMoldGroups_(sheet);          // bands every column, incl. Mold ID + Description, by mold
  greyOutDescriptionColumn_(sheet); // Description column always wins back to grey
  populateDescriptions_(sheet, DATA_START_ROW, lastRow); // values + red flag on unmatched Mold IDs
}

function getMoldDescriptionMap_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ALL_MOLDS_SHEET_NAME);
  if (!sheet) throw new Error('"' + ALL_MOLDS_SHEET_NAME + '" tab not found.');
  const headerRow = findHeaderRowInColA_(sheet, MOLD_ID_HEADER);
  const idCol = findHeaderCol_(sheet, headerRow, MOLD_ID_HEADER);
  const descCol = findHeaderCol_(sheet, headerRow, MOLD_DESC_HEADER);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow > headerRow) {
    const ids = sheet.getRange(headerRow + 1, idCol, lastRow - headerRow, 1).getValues();
    const descs = sheet.getRange(headerRow + 1, descCol, lastRow - headerRow, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0] || '').trim();
      if (id && !map.hasOwnProperty(id)) map[id] = String(descs[i][0] || '').trim();
    }
  }
  return map;
}

function populateDescriptions_(sheet, startRow, endRow) {
  const moldIdCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_ID_HEADER);
  const descCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_DESC_HEADER);
  if (!moldIdCol || !descCol) return;

  const map = getMoldDescriptionMap_();
  const numRows = endRow - startRow + 1;
  const idRange = sheet.getRange(startRow, moldIdCol, numRows, 1);
  const ids = idRange.getValues();
  const idBgs = idRange.getBackgrounds(); // preserve mold-group band already applied
  const descValues = [];

  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0] || '').trim();
    if (!id) { descValues.push(['']); continue; }
    if (map.hasOwnProperty(id)) {
      descValues.push([map[id]]);
    } else {
      descValues.push(['']);
      idBgs[i][0] = NOT_FOUND_BG; // flag immediately — Mold ID not in All Molds List
    }
  }

  sheet.getRange(startRow, descCol, numRows, 1).setValues(descValues);
  idRange.setBackgrounds(idBgs);
}

function greyOutDescriptionColumn_(sheet) {
  const descCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_DESC_HEADER);
  if (!descCol) return;
  const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  sheet.getRange(DATA_START_ROW, descCol, lastRow - DATA_START_ROW + 1, 1).setBackground(DESC_GREY);
}

function shadeMoldGroups_(sheet) {
  const moldIdCol = findHeaderCol_(sheet, HEADER_ROW, MOLD_ID_HEADER);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!moldIdCol || lastRow < DATA_START_ROW) return;

  const ids = sheet.getRange(DATA_START_ROW, moldIdCol, lastRow - DATA_START_ROW + 1, 1).getValues();
  const bgRows = [];
  let toggle = false;
  let prevId = null;
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0] || '').trim();
    if (id !== prevId) { toggle = !toggle; prevId = id; }
    const color = toggle ? SHADE_B : SHADE_A;
    const row = [];
    for (let c = 0; c < lastCol; c++) row.push(color);
    bgRows.push(row);
  }
  sheet.getRange(DATA_START_ROW, 1, ids.length, lastCol).setBackgrounds(bgRows);
}

// ================= HEADER HELPERS =================
function findHeaderCol_(sheet, headerRow, label) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === label) return i + 1;
  }
  return null;
}

function findHeaderRowInColA_(sheet, label) {
  const lastRow = Math.min(sheet.getLastRow(), 10);
  for (let r = 1; r <= lastRow; r++) {
    if (String(sheet.getRange(r, 1).getValue()).trim() === label) return r;
  }
  throw new Error('Could not find header "' + label + '" in column A of "' + sheet.getName() + '".');
}
