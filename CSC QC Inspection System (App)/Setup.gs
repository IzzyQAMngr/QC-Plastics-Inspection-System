/*************************************************************
 * ONE-TIME SETUP — run once from the Apps Script editor
 * (select oneTimeSetup, click Run) after clasp push.
 *************************************************************/
function oneTimeSetup() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('DATA_SPREADSHEET_ID', ss.getId());

  ensureSheetWithHeaders_(ss, DROPFREEZE_LOG_SHEET_NAME, DROPFREEZE_LOG_HEADERS);
  const inProcessSheet = ss.getSheetByName(INPROCESS_LOG_SHEET_NAME);
  ensureColumnExists_(inProcessSheet, 'BatchID');
  ensureColumnExists_(inProcessSheet, 'Item No.');
  ensureColumnExists_(inProcessSheet, 'Item Description');
  ensureColumnExists_(inProcessSheet, 'Release Decision');
  ensureColumnExists_(inProcessSheet, 'Justification');

  const messages = ['DATA_SPREADSHEET_ID set to this spreadsheet.'];
  if (!ss.getSheetByName(SETTINGS_SHEET_NAME)) {
    messages.push('⚠️ No "Settings" sheet found — Master Register ID / notification emails / mold & personnel lists will need to be added there.');
  }
  Logger.log(messages.join(' '));
  try { SpreadsheetApp.getUi().alert(messages.join('\n')); } catch (e) { /* running headless */ }
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Safe to run only while Drop Freeze Test Data has no real rows yet — rewrites its header
 * row to the current DROPFREEZE_LOG_HEADERS whenever that list changes (e.g. adding
 * ProductType/ItemNo/ItemDescription). Throws once real rows exist; migrate by hand after
 * that point instead. Run once after any header-list change.
 */
function migrateDropFreezeHeaders() {
  const sheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getActive().toast(DROPFREEZE_LOG_SHEET_NAME + ' sheet not found.'); return; }
  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    throw new Error(DROPFREEZE_LOG_SHEET_NAME + ' has ' + dataRows + ' real row(s) already — rewriting headers would misalign them. Migrate manually instead.');
  }
  sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).clearContent();
  sheet.getRange(1, 1, 1, DROPFREEZE_LOG_HEADERS.length).setValues([DROPFREEZE_LOG_HEADERS]).setFontWeight('bold');
  SpreadsheetApp.getActive().toast(DROPFREEZE_LOG_SHEET_NAME + ' headers updated: ' + DROPFREEZE_LOG_HEADERS.join(', '));
}

function ensureColumnExists_(sheet, headerName) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(headerName) === -1) {
    sheet.getRange(1, lastCol + 1).setValue(headerName).setFontWeight('bold');
  }
}

/*************************************************************
 * SETTINGS REDESIGN — run once from the Apps Script editor.
 * Non-destructive: renames the current scattered "Settings" sheet
 * to "Settings (old)" (untouched, kept as a backup) and builds a
 * clean new "Settings" sheet from the values found in it. Delete
 * "Settings (old)" yourself once you've confirmed the new one looks right.
 *************************************************************/
function migrateSettingsTab() {
  const ss = getDb_();
  const oldSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!oldSheet) throw new Error('"' + SETTINGS_SHEET_NAME + '" sheet not found.');
  if (ss.getSheetByName(SETTINGS_SHEET_NAME + ' (old)')) {
    throw new Error('"' + SETTINGS_SHEET_NAME + ' (old)" already exists — migration already run once. Delete or rename it first to re-run.');
  }

  const lastRow = oldSheet.getLastRow();
  const lastCol = oldSheet.getLastColumn();
  const values = oldSheet.getRange(1, 1, lastRow, lastCol).getValues();

  function findLabel_(label) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (String(values[r][c]).trim() === label) return { r: r, c: c };
      }
    }
    return null;
  }

  function valueRightOf_(label) {
    const pos = findLabel_(label);
    if (!pos) return '';
    const v = values[pos.r][pos.c + 1];
    return v === undefined ? '' : String(v).trim();
  }

  function columnListBelow_(pos, maxScan) {
    if (!pos) return [];
    const out = [];
    let blanks = 0;
    for (let i = 1; i <= maxScan && blanks < 6; i++) {
      const row = values[pos.r + i];
      const v = row ? row[pos.c] : undefined;
      const s = v === undefined ? '' : String(v).trim();
      if (s === '') { blanks++; continue; }
      blanks = 0;
      out.push(s);
    }
    return out;
  }

  function tableRowsBelow_(pos, colOffsets, maxScan) {
    if (!pos) return [];
    const out = [];
    let blanks = 0;
    for (let i = 1; i <= maxScan && blanks < 8; i++) {
      const row = values[pos.r + i];
      const cells = colOffsets.map(function (off) {
        const v = row ? row[pos.c + off] : undefined;
        return v === undefined ? '' : String(v).trim();
      });
      if (cells.every(function (c) { return c === ''; })) { blanks++; continue; }
      blanks = 0;
      out.push(cells);
    }
    return out;
  }

  // ---- Pull everything worth keeping out of the old scattered layout ----
  const notificationEmails = valueRightOf_('Notification Emails (comma-separated):');
  const inspectors  = columnListBelow_(findLabel_('Inspected By'), 20);
  const foremen     = columnListBelow_(findLabel_('Shift Foreman'), 20);
  const shifts      = columnListBelow_(findLabel_('Shift'), 20);
  const passFailVals = columnListBelow_(findLabel_('Date Code/Nesting/Cover/Gauge'), 20);
  // Item List and Version History turned out to live in their own tabs ("All Items List",
  // "Version History"), not embedded in Settings — nothing to pull for those here.

  // ---- Rename old sheet as a backup, build a clean new one in its place ----
  const oldIndex = oldSheet.getIndex();
  oldSheet.setName(SETTINGS_SHEET_NAME + ' (old)');
  const sheet = ss.insertSheet(SETTINGS_SHEET_NAME, oldIndex - 1);

  function sectionHeader_(row, text, span) {
    const range = sheet.getRange(row, 1, 1, span || 4);
    range.merge();
    range.setValue(text).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  }

  sheet.getRange('A1').setValue('CSC QC Inspection System — Settings').setFontWeight('bold').setFontSize(14);

  let r = 3;
  sectionHeader_(r, 'MASTER REGISTER', 2); r++;
  sheet.getRange(r, 1).setValue('Master Register ID (Spec Register):');
  sheet.getRange(r, 2).setValue(SPEC_REGISTER_ID);
  r += 2;

  sectionHeader_(r, 'NOTIFICATIONS', 2); r++;
  sheet.getRange(r, 1).setValue('Notification Emails (comma-separated):');
  sheet.getRange(r, 2).setValue(notificationEmails);
  r += 2;

  sectionHeader_(r, 'HEADER DROPDOWNS', 4); r++;
  sheet.getRange(r, 1, 1, 4).setValues([['Inspected By', 'Shift Foreman', 'Shift', 'Pass / Fail / N/A']]).setFontWeight('bold');
  r++;
  const maxLen = Math.max(inspectors.length, foremen.length, shifts.length, passFailVals.length, 1);
  const dropdownRows = [];
  for (let i = 0; i < maxLen; i++) {
    dropdownRows.push([inspectors[i] || '', foremen[i] || '', shifts[i] || '', passFailVals[i] || '']);
  }
  if (dropdownRows.length) sheet.getRange(r, 1, dropdownRows.length, 4).setValues(dropdownRows);

  sheet.setColumnWidths(1, 4, 220);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getActive().toast(
    'Settings redesigned: ' + inspectors.length + ' inspectors, ' + foremen.length + ' foremen copied over. ' +
    'Old layout kept as "Settings (old)" — delete it yourself once you\'ve confirmed the new one looks right.',
    'Settings Redesigned', 10
  );
}

/**
 * One-time: creates the "Line Configuration" sheet (Department | Line # | Equipment Code |
 * Equipment Description — one row per piece of equipment on a line) and seeds it with
 * Plastics' existing Line 1-10 (no equipment breakdown for Plastics, so those columns stay
 * blank). Metals rows are added by hand afterward. No-ops if the sheet already has data,
 * so it's safe to run again.
 */
function setupLineConfiguration() {
  const ss = getDb_();
  const sheet = ensureSheetWithHeaders_(ss, LINE_CONFIG_SHEET_NAME, ['Department', 'Line #', 'Equipment Code', 'Equipment Description']);
  if (sheet.getLastRow() > 1) { SpreadsheetApp.getActive().toast('Line Configuration already has data — leaving it as-is.'); return; }
  const rows = [];
  for (let i = 1; i <= 10; i++) rows.push(['Plastics', String(i), '', '']);
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  SpreadsheetApp.getActive().toast('Line Configuration created and seeded with Plastics Line 1-10. Add Metals rows by hand.');
}

/** One-time: creates the "Runs - Metals" sheet with headers, mirroring the Plastics
 *  "Runs - Plastics" sheet but with Material Lot / Size ID / Can Description in place of
 *  Resin Lot / Mold ID / Mold Description. No-ops if the sheet already exists. */
function setupMetalsRunsSheet() {
  const ss = getDb_();
  ensureSheetWithHeaders_(ss, RUNS_SHEET_NAME_METALS, METALS_RUNS_HEADERS);
  SpreadsheetApp.getActive().toast('"' + RUNS_SHEET_NAME_METALS + '" is ready.');
}

/** One-time: creates the "QC Inspection Data- Metals Ends" flat log sheet with headers.
 *  No-ops if the sheet already exists. Run once from the Apps Script editor after clasp push. */
function setupMetalsEndsLogSheet() {
  const ss = getDb_();
  ensureSheetWithHeaders_(ss, METALS_ENDS_LOG_SHEET_NAME, METALS_ENDS_LOG_HEADERS);
  SpreadsheetApp.getActive().toast('"' + METALS_ENDS_LOG_SHEET_NAME + '" is ready.');
}

/**
 * ONE-TIME: imports historical Drop Freeze records from the old system's spreadsheet
 * ("Test Data" tab) into this system's Drop Freeze Test Data log, mapping the old column
 * names onto the current DROPFREEZE_LOG_HEADERS schema. Fields the old system never tracked
 * (Run ID, Customer Name, Mold Description, Product Type, Item No, Item Description) come in
 * blank — those are populated going forward by the Run-driven Drop Freeze form.
 * Refuses to run if the destination sheet already has real rows, to avoid double-importing.
 * Run once from the Apps Script editor (select migrateOldDropFreezeData, click Run).
 */
function migrateOldDropFreezeData() {
  const OLD_SPREADSHEET_ID = '1kxqvqPVA7l5-EgkyqyXqf4xmOI7oukNfTjnq_GTkrAo';
  const OLD_SHEET_NAME = 'Test Data';

  const newSheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  if (!newSheet) throw new Error('"' + DROPFREEZE_LOG_SHEET_NAME + '" sheet not found — run oneTimeSetup first.');
  const existingDataRows = newSheet.getLastRow() - 1;
  if (existingDataRows > 0) {
    throw new Error('"' + DROPFREEZE_LOG_SHEET_NAME + '" already has ' + existingDataRows + ' row(s) — refusing to ' +
      'risk a duplicate import. Clear the data rows (keep the header) first if you want to re-run this.');
  }

  const oldSheet = SpreadsheetApp.openById(OLD_SPREADSHEET_ID).getSheetByName(OLD_SHEET_NAME);
  if (!oldSheet) throw new Error('"' + OLD_SHEET_NAME + '" sheet not found in the old spreadsheet.');
  const oldRows = readSheetObjects_(oldSheet);

  const mapped = oldRows.map(r => ({
    'RecordKey': r['QC Record #'] || '',
    'LineItem': r['Line Item No.'] || '',
    'Status': r['Status'] || '',
    'Created': r['Created TS'] || '',
    'Updated': r['Last Updated TS'] || '',
    'Line #': r['Line #'] || '',
    'Shift': r['Shift'] || '',
    'Mold ID': r['Tool Code'] || '',
    'Resin Lot': r['Resin ID'] || '',
    'Cavity': r['Cavity'] || '',
    'Test Name': r['Test Type'] || '',
    'DateOfMfg': r['Date of Mfg'] || '',
    'TestDate': r['Test Date'] || '',
    'TestedBy': r['Tested By'] || '',
    'FreezerTemp': r['Freezer Temp. ( deg f)'] || '',
    'DropHeight': r['Drop Height'] || '',
    'DropAngle': r['Drop Angle'] || '',
    'Result': r['Pass/ Fail/In Freezer'] || '',
    'FailureDescription': r['Failure Description (Indicate wall side, any other observations)'] || '',
    'Notes': r['Other Notes/ Observations'] || '',
    'Month': r['Month'] || '',
    'Year': r['Year'] || '',
  }));

  appendObjectsAsRows_(newSheet, mapped);
  SpreadsheetApp.getActive().toast('Imported ' + mapped.length + ' Drop Freeze records from the old system into "' + DROPFREEZE_LOG_SHEET_NAME + '".');
}

/** One-off cleanup: removes the empty "ITEM LIST" placeholder section a previous
 *  run of migrateSettingsTab left behind before Item List was found in its own
 *  "All Items List" tab instead. Safe to run once; no-ops if already clean. */
function removeEmptyItemListSectionFromSettings() {
  const sheet = getDb_().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  let startRow = -1;
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim().indexOf('ITEM LIST') === 0) { startRow = r + 1; break; }
  }
  if (startRow === -1) { SpreadsheetApp.getActive().toast('No leftover ITEM LIST section found — already clean.'); return; }
  sheet.deleteRows(startRow, lastRow - startRow + 1);
  SpreadsheetApp.getActive().toast('Removed empty ITEM LIST placeholder from Settings.');
}

/**
 * ONE-TIME: adds the Release Decision / Justification columns to the Plastics QC Inspection
 * Data log (if missing — oneTimeSetup also does this going forward), then migrates every
 * already-saved deviation sign-off that's still riding along as a fake "Deviation" row
 * (Test Type = "Deviation", Characteristic Name = "Release Decision"/"Justification") onto
 * the real columns for its matching sample (same QC Record # + Inspection ID), and removes
 * those pseudo-rows. Safe to re-run — a clean run after the first just no-ops.
 * Run once from the Apps Script editor (select migrateInProcessReleaseDecision, click Run).
 */
function migrateInProcessReleaseDecision() {
  const sheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + INPROCESS_LOG_SHEET_NAME + '" sheet not found.');
  ensureColumnExists_(sheet, 'Release Decision');
  ensureColumnExists_(sheet, 'Justification');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) { SpreadsheetApp.getActive().toast('No data rows — nothing to migrate.'); return; }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rdCol = headers.indexOf('Release Decision');
  const jCol = headers.indexOf('Justification');
  const recordCol = headers.indexOf('QC Record #');
  const idCol = headers.indexOf('Inspection ID');
  const typeCol = headers.indexOf('Test Type');
  const charCol = headers.indexOf('Characteristic Name');
  const valCol = headers.indexOf('Actual Value');

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const bySample = {}; // 'QC Record #|Inspection ID' -> {decision, justification}
  const pseudoRows = []; // 1-based sheet row numbers to remove afterward
  data.forEach((row, i) => {
    if (String(row[typeCol]) !== 'Deviation') return;
    const charName = String(row[charCol]);
    if (charName !== 'Release Decision' && charName !== 'Justification') return;
    const key = row[recordCol] + '|' + row[idCol];
    bySample[key] = bySample[key] || {};
    bySample[key][charName === 'Release Decision' ? 'decision' : 'justification'] = row[valCol];
    pseudoRows.push(i + 2);
  });

  if (pseudoRows.length === 0) {
    SpreadsheetApp.getActive().toast('No legacy Deviation pseudo-rows found — nothing to migrate.');
    return;
  }

  const rdValues = [];
  const jValues = [];
  let stamped = 0;
  data.forEach(row => {
    const key = row[recordCol] + '|' + row[idCol];
    const info = bySample[key];
    if (info && String(row[typeCol]) !== 'Deviation') {
      rdValues.push([info.decision !== undefined ? info.decision : row[rdCol]]);
      jValues.push([info.justification !== undefined ? info.justification : row[jCol]]);
      stamped++;
    } else {
      rdValues.push([row[rdCol]]);
      jValues.push([row[jCol]]);
    }
  });
  sheet.getRange(2, rdCol + 1, rdValues.length, 1).setValues(rdValues);
  sheet.getRange(2, jCol + 1, jValues.length, 1).setValues(jValues);

  // Bottom-up so row numbers stay valid as each delete shifts everything below it up.
  pseudoRows.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

  const msg = 'Migrated ' + Object.keys(bySample).length + ' sample(s): stamped ' + stamped +
    ' data row(s), removed ' + pseudoRows.length + ' legacy pseudo-row(s).';
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg); } catch (e) { /* running headless */ }
}

/**
 * ONE-TIME: replaces the Plastics QC Inspection Data log's legacy Month/Year
 * ARRAYFORMULAs (=ARRAYFORMULA(IF(D2:D="","",MONTH(D2:D))) and its YEAR() counterpart, both
 * living only in row 2) with plain per-row values computed from Inspection Date — matching
 * what saveInProcessInspection() now writes in code for every new row (see InProcess.gs). The
 * array formula silently stopped covering the whole column the moment any row below it got a
 * literal value from the code (Sheets refuses to let an ARRAYFORMULA's spill range overwrite
 * existing data), which is why Month/Year currently read blank for most existing rows — this
 * backfills every row at once and leaves no formula behind to break again.
 * Run once from the Apps Script editor (select fixInProcessMonthYear, click Run).
 */
function fixInProcessMonthYear() {
  const sheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  if (!sheet) throw new Error('"' + INPROCESS_LOG_SHEET_NAME + '" sheet not found.');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) { SpreadsheetApp.getActive().toast('No data rows — nothing to fix.'); return; }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const dateCol = headers.indexOf('Inspection Date');
  const monthCol = headers.indexOf('Month');
  const yearCol = headers.indexOf('Year');
  if (dateCol < 0 || monthCol < 0 || yearCol < 0) {
    throw new Error('Missing one of Inspection Date/Month/Year columns.');
  }

  const dates = sheet.getRange(2, dateCol + 1, lastRow - 1, 1).getValues();
  const monthValues = [];
  const yearValues = [];
  let fixed = 0;
  dates.forEach(([raw]) => {
    let y = '', m = '';
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      y = raw.getFullYear(); m = raw.getMonth() + 1;
    } else {
      const s = String(raw || '').trim();
      const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/); // yyyy-mm-dd
      const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // m/d/yyyy
      if (iso) { y = Number(iso[1]); m = Number(iso[2]); }
      else if (slash) { y = Number(slash[3]); m = Number(slash[1]); }
    }
    if (y !== '') fixed++;
    monthValues.push([m]);
    yearValues.push([y]);
  });

  // setValues overwrites row 2's ARRAYFORMULA cell too — no formula survives this.
  sheet.getRange(2, monthCol + 1, monthValues.length, 1).setValues(monthValues);
  sheet.getRange(2, yearCol + 1, yearValues.length, 1).setValues(yearValues);

  const msg = 'Recomputed Month/Year for ' + fixed + ' of ' + dates.length + ' row(s) from Inspection Date.';
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg); } catch (e) { /* running headless */ }
}
