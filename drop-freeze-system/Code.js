/********************************************************
 * QC FORM → TEST DATA (SAVE on P3) + LOAD on J3 + CLEAR on P4
 * - Dynamic grid (B5:Q...) with headers in row 5
 * - P3 checked: saves/updates ONE packet record with multiple line-items into "Test Data"
 * - J3 dropdown: loads selected packet back into grid
 * - AFTER SAVE: clears the form (grid + D3 + J3), keeps formulas
 *
 * + CLEAR FORM (NEW RECORD) ON P4
 * - P4 checked: clears grid data (keeps formulas) + clears D3 + clears J3
 * - Resets P4 to FALSE after clearing
 *
 * + GRAY OUT GRID WHEN RECORD IS LOADED
 * - When J3 loads a record, ONLY the data-entry rows (row 6+) are shaded.
 * - Header row 5 is NOT changed (keeps your colored headers).
 * - When you Clear Form (P4) or after Save (auto-clear), grid shading is restored.
 *
 * + FAIL EMAIL ALERT (BASED ON SAVED DATA)
 * - QC Form Result column is O (row 6+)
 * - That maps to Test Data column S (row 2+)
 * - Emails you after SAVE if any saved line-item has FAIL in that field
 * - De-duped per Record# + Line Item (clears if no longer FAIL)
 *
 * STATUS LOGIC:
 * - Pass / Fail / Inconclusive → Status = COMPLETE
 * - Blank result → Status = OPEN
 * - VOID = manually set by admin (excluded from all reporting)
 *
 * DROPDOWN:
 * - J3 validation list rebuilt after every save/clear — OPEN records only
 *
 * FORM DESIGN PERSISTENCE:
 * - Colors AND fonts restored after every save, clear, and load
 * - Load state uses soft gray tint
 * - redesignQCForm() re-runs from QC Tools menu anytime
 *
 * DASHBOARD SYSTEM (v9):
 * - 📊 DASHBOARD — KPIs + summary tables (live COUNTIFS)
 * - ⏳ In Progress — FILTER-based live aging tracker (fixed column mapping)
 * - 🔍 MOLD ANALYSIS — Drill-down by mold with month/year, test type + temps,
 *   angle breakdowns, and sparkline trend visuals
 *
 * + MASTER REGISTER TOOL CODE LOADER
 * - Pulls Tool Codes from Master Register → All Molds tab → Col C
 * - Writes them into Data Validation Lists!A3 for the F6 dropdown
 * - Run via QC Tools menu → Load Tool Codes from Master Register
 ********************************************************/

// ================= CONFIG =================
const FORM_SHEET_NAME = 'QC Form';
const LOG_SHEET_NAME  = 'Test Data';

const SUBMIT_CELL_A1     = 'P3';
const CLEAR_CELL_A1      = 'P4';
const DROPDOWN_CELL_A1   = 'J3';
const RECORD_KEY_CELL_A1 = 'D3';

const HEADER_ROW     = 5;
const DATA_START_ROW = 6;
const START_COL      = 2;
const END_COL        = 17;

const RESULT_REL_IDX = 13;

const LOG_COL_RECORDKEY  = 1;
const LOG_COL_LINEITEM   = 2;
const LOG_COL_STATUS     = 3;
const LOG_COL_CREATED    = 4;
const LOG_COL_UPDATED    = 5;
const LOG_COL_DATA_START = 6;

// ── Master Register link for tool code dropdown ──
const DF_SETTINGS_SHEET = "Settings";
const DF_MR_ID_CELL     = "B5";

// ================= FORM DESIGN COLORS =================
const FORM_DESIGN = {
  sampleOdd:       '#FFFFFF',
  sampleEven:      '#FFFDF5',
  resultOdd:       '#FFFFFF',
  resultEven:      '#F5FBFF',
  lineOdd:         '#FFFEF5',
  lineEven:        '#FFF8E6',
  passfailOdd:     '#F5FCFF',
  passfailEven:    '#EFF9FF',
  sampleGrayOdd:   '#F4F2EC',
  sampleGrayEven:  '#EEECE4',
  resultGrayOdd:   '#EBF2F5',
  resultGrayEven:  '#E4EEF3',
  lineGrayOdd:     '#F0EDDF',
  lineGrayEven:    '#EAE5D5',
  passfailGrayOdd: '#E6EFF3',
  passfailGrayEven:'#DDE9EF',
  margin:          '#1F3864',
};

// ================= FAIL EMAIL CONFIG =================
const FAIL_EMAIL_TO = "izuniga@cscmfg.com";
const FAIL_EMAIL_CC = '';

function isFailValue_(val) {
  const s = String(val || '').trim();
  if (!s) return false;
  const up = s.toUpperCase();
  return up.includes('FAIL') || s.includes('❌');
}


// ============== INSTALLABLE TRIGGER ENTRY ==============
function onEditSubmitTrigger(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== FORM_SHEET_NAME) return;

  const a1 = range.getA1Notation();

  if (a1 === CLEAR_CELL_A1) {
    if (range.getValue() !== true) return;
    SpreadsheetApp.getActive().toast('Clearing QC form for NEW record...', 'Clearing', -1);
    try {
      clearQcFormForNewRecordCore_();
      SpreadsheetApp.getActive().toast('✅ Form cleared. Next save will create a NEW record.', 'Cleared', 5);
    } catch (err) {
      SpreadsheetApp.getActive().toast('❌ Clear failed.', 'ERROR', 8);
      throw err;
    } finally {
      try { sheet.getRange(CLEAR_CELL_A1).setValue(false); } catch (_) {}
    }
    return;
  }

  if (a1 === DROPDOWN_CELL_A1) {
    if (String(range.getDisplayValue() || '').trim() === '') return;
    SpreadsheetApp.getActive().toast('Loading QC record...', 'Loading', -1);
    loadQcPacketFromDropdownCore_();
    return;
  }

  if (a1 !== SUBMIT_CELL_A1) return;
  if (range.getValue() !== true) return;

  SpreadsheetApp.getActive().toast('Saving QC record to Test Data...', 'Submitting', -1);

  try {
    submitQCCore_();
    try { flashGrid_(600); } catch (_) {}
    try { clearQcFormAfterSave_(); } catch (clearErr) {
      SpreadsheetApp.getActive().toast(
        'Saved OK, but form clear failed: ' + (clearErr && clearErr.message ? clearErr.message : clearErr),
        'Clear Warning', 10
      );
    }
    SpreadsheetApp.getActive().toast('✅ QC record saved & form cleared.', 'SUCCESS', 5);
  } catch (err) {
    SpreadsheetApp.getActive().toast('QC submit failed.', 'ERROR', 8);
    throw err;
  } finally {
    try { sheet.getRange(SUBMIT_CELL_A1).setValue(false); } catch (_) {}
  }
}


// ============== CORE CLEAR LOGIC (P4) ==============
function clearQcFormForNewRecordCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_SHEET_NAME);
  if (!sheet) return;

  clearGridDataOnly_(sheet);
  restoreDataRowDesign_(sheet, false);
  sheet.getRange(RECORD_KEY_CELL_A1).clearContent();
  sheet.getRange(DROPDOWN_CELL_A1).clearContent();

  try { refreshDropdownValidation_(); } catch (_) {}
}


// ============== CORE LOAD LOGIC (J3) ==============
function loadQcPacketFromDropdownCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
  const logSheet  = ss.getSheetByName(LOG_SHEET_NAME);

  if (!formSheet || !logSheet) {
    SpreadsheetApp.getUi().alert('Missing required sheet(s). Check sheet names.');
    return;
  }

  const recordKey = String(formSheet.getRange(DROPDOWN_CELL_A1).getDisplayValue() || '').trim();
  if (!recordKey) return;

  if (!/^QC-\d{8}-\d{3}$/.test(recordKey)) {
    SpreadsheetApp.getUi().alert('Invalid record selected in ' + DROPDOWN_CELL_A1 + ': ' + recordKey);
    return;
  }

  const rows = getAllLogRowsForRecord_(logSheet, recordKey);
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('No line items found for record: ' + recordKey);
    return;
  }

  const width  = END_COL - START_COL + 1;
  const qcData = rows.map(r => r.slice(LOG_COL_DATA_START - 1, LOG_COL_DATA_START - 1 + width));

  clearGridDataOnly_(formSheet);

  const loadRange  = formSheet.getRange(DATA_START_ROW, START_COL, qcData.length, width);
  const savedRules = loadRange.getDataValidations();
  loadRange.clearDataValidations();
  loadRange.setValues(qcData);
  loadRange.setDataValidations(savedRules);

  formSheet.getRange(RECORD_KEY_CELL_A1).setValue(recordKey);
  restoreDataRowDesign_(formSheet, true);

  SpreadsheetApp.getActive().toast('✅ Loaded ' + qcData.length + ' line item(s).', 'Loaded', 5);
}


// ============== CORE SAVE LOGIC (P3) ==============
function submitQCCore_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
    const logSheet  = ss.getSheetByName(LOG_SHEET_NAME);

    if (!formSheet || !logSheet) {
      SpreadsheetApp.getUi().alert('Missing required sheet(s). Check sheet names.');
      return;
    }

    const width = END_COL - START_COL + 1;

    const toolRelIdx = findHeaderIndex_(formSheet, HEADER_ROW, START_COL, width, /tool\s*code/i);
    if (toolRelIdx === -1) {
      SpreadsheetApp.getUi().alert(
        'I could not find a "Tool Code" header in row ' + HEADER_ROW +
        ' (within B:Q). Please make sure the header text contains "Tool Code".'
      );
      return;
    }

    const toolAbsCol  = START_COL + toolRelIdx;
    const lastDataRow = findLastDataRowByToolCode_(formSheet, DATA_START_ROW, toolAbsCol);
    if (lastDataRow < DATA_START_ROW) {
      SpreadsheetApp.getUi().alert('No line items found (Tool Code column is blank).');
      return;
    }

    const numRows    = lastDataRow - HEADER_ROW + 1;
    const gridRange  = formSheet.getRange(HEADER_ROW, START_COL, numRows, width);
    const gridValues = gridRange.getDisplayValues();
    const dataRows   = gridValues.slice(1);

    const active = [];
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const toolCode = (row[toolRelIdx] || '').toString().trim();
      if (!toolCode) continue;
      active.push(row);
    }

    if (active.length === 0) {
      SpreadsheetApp.getUi().alert('No active line items found (Tool Code cells are blank).');
      return;
    }

    const recordCell = formSheet.getRange(RECORD_KEY_CELL_A1);
    let recordKey = String(recordCell.getDisplayValue() || '').trim();

    if (recordKey && !/^QC-\d{8}-\d{3}$/.test(recordKey)) {
      SpreadsheetApp.getUi().alert(
        'Invalid QC Record # in ' + RECORD_KEY_CELL_A1 + ': "' + recordKey + '"\n' +
        'That cell must contain QC-YYYYMMDD-### or be blank.'
      );
      return;
    }

    const firstDateOfMfgDisplay = (active[0][1] || '').toString().trim();

    if (!recordKey) {
      recordKey = makeDailyRecordKey_(logSheet, firstDateOfMfgDisplay, ss);
      recordCell.setValue(recordKey);
    }

    const updatedAtNow = new Date();
    const existingInfo = findExistingRecordRows_(logSheet, recordKey);
    let createdTS = new Date();

    if (existingInfo.found) {
      const maybeCreated = logSheet.getRange(existingInfo.firstRow, LOG_COL_CREATED).getValue();
      if (maybeCreated instanceof Date) createdTS = maybeCreated;
      const sortedRows  = existingInfo.rows.slice().sort((a, b) => a - b);
      const firstRow    = sortedRows[0];
      const numToDelete = sortedRows.length;
      logSheet.deleteRows(firstRow, numToDelete);
    }

    const rowsToAppend = active.map((row, idx) => {
      const resultRaw  = (row[RESULT_REL_IDX] || '').toString().trim();
      const normalized = resultRaw.replace(/[^\w\s]/g, '').toUpperCase();

      let status = 'OPEN';
      if (
        normalized.includes('PASS') ||
        normalized.includes('FAIL') ||
        normalized.includes('INCONCLUSIVE')
      ) {
        status = 'COMPLETE';
      }

      return [recordKey, idx + 1, status, createdTS, updatedAtNow, ...row];
    });

    logSheet
      .getRange(logSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);

    try { sendFailEmailsForSavedPacket_(recordKey, active, ss); } catch (_) {}

  } finally {
    try { refreshDropdownValidation_(); } catch (_) {}
    lock.releaseLock();
  }
}


// ============== FAIL EMAILS (AFTER SAVE) ==============
function sendFailEmailsForSavedPacket_(recordKey, activeRows, ss) {
  const props = PropertiesService.getScriptProperties();
  const ssId  = ss.getId();

  const formSheet     = ss.getSheetByName(FORM_SHEET_NAME);
  const testDataSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const testDataGid   = testDataSheet.getSheetId();

  const width   = END_COL - START_COL + 1;
  const headers = formSheet
    .getRange(HEADER_ROW, START_COL, 1, width)
    .getDisplayValues()[0];

  for (let i = 0; i < activeRows.length; i++) {
    const lineItem  = i + 1;
    const resultVal = (activeRows[i][RESULT_REL_IDX] || '').toString().trim();
    const dedupeKey = `FAIL_SAVED__${ssId}__${recordKey}__LINE_${lineItem}`;

    if (!isFailValue_(resultVal)) {
      props.deleteProperty(dedupeKey);
      continue;
    }

    if (props.getProperty(dedupeKey)) continue;

    const labeledLines = [];
    for (let c = 0; c < headers.length; c++) {
      const header = headers[c];
      const value  = activeRows[i][c];
      if (header && String(header).trim() && String(value || '').trim() !== '') {
        labeledLines.push(`${header}: ${value}`);
      }
    }

    const testDataRow = lineItem + 1;
    const deepLink    = ss.getUrl() + `#gid=${testDataGid}&range=S${testDataRow}`;
    const subject     = `QC Freeze Drop Test FAIL Saved — ${recordKey} (Line ${lineItem})`;
    const body =
`A FAIL result was saved to Test Data (column S).

Record #: ${recordKey}
Line Item #: ${lineItem}

--------------------------------
LINE ITEM DETAILS
--------------------------------
${labeledLines.join('\n')}

--------------------------------
Open Test Data (FAIL highlighted):
${deepLink}
`;

    MailApp.sendEmail({ to: FAIL_EMAIL_TO, cc: FAIL_EMAIL_CC, subject, body });
    props.setProperty(dedupeKey, new Date().toISOString());
  }
}


// ============== CLEAR FORM AFTER SAVE ==============
function clearQcFormAfterSave_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_SHEET_NAME);
  clearGridDataOnly_(sheet);
  restoreDataRowDesign_(sheet, false);
  sheet.getRange(RECORD_KEY_CELL_A1).clearContent();
  sheet.getRange(DROPDOWN_CELL_A1).clearContent();
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  DESIGN-AWARE ROW COLORING                                              ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function restoreDataRowDesign_(formSheet, grayMode) {
  const D       = FORM_DESIGN;
  const lastRow = Math.max(formSheet.getLastRow(), DATA_START_ROW);
  const maxRow  = Math.max(lastRow, 65);
  const numRows = maxRow - DATA_START_ROW + 1;

  const sampleBgs = [];
  const resultBgs = [];

  for (let r = DATA_START_ROW; r <= maxRow; r++) {
    const isEven = (r % 2 === 0);

    const sRow = [];
    for (let c = 2; c <= 9; c++) {
      if (c === 2) {
        sRow.push(grayMode
          ? (isEven ? D.lineGrayEven  : D.lineGrayOdd)
          : (isEven ? D.lineEven      : D.lineOdd));
      } else {
        sRow.push(grayMode
          ? (isEven ? D.sampleGrayEven : D.sampleGrayOdd)
          : (isEven ? D.sampleEven     : D.sampleOdd));
      }
    }
    sampleBgs.push(sRow);

    const rRow = [];
    for (let c = 10; c <= 17; c++) {
      if (c === 15) {
        rRow.push(grayMode
          ? (isEven ? D.passfailGrayEven : D.passfailGrayOdd)
          : (isEven ? D.passfailEven      : D.passfailOdd));
      } else {
        rRow.push(grayMode
          ? (isEven ? D.resultGrayEven : D.resultGrayOdd)
          : (isEven ? D.resultEven     : D.resultOdd));
      }
    }
    resultBgs.push(rRow);
  }

  formSheet.getRange(DATA_START_ROW, 2,  numRows, 8).setBackgrounds(sampleBgs);
  formSheet.getRange(DATA_START_ROW, 10, numRows, 8).setBackgrounds(resultBgs);
  formSheet.getRange(DATA_START_ROW, 1,  numRows, 1).setBackground(D.margin);
  formSheet.getRange(DATA_START_ROW, 18, numRows, 1).setBackground(D.margin);

  const fullDataRange = formSheet.getRange(DATA_START_ROW, START_COL, numRows, END_COL - START_COL + 1);
  fullDataRange
    .setFontFamily('Calibri')
    .setFontSize(10)
    .setFontWeight('normal')
    .setFontStyle('normal')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  formSheet.getRange(DATA_START_ROW, 2, numRows, 1)
    .setFontSize(10)
    .setFontColor('#888888')
    .setHorizontalAlignment('center');

  formSheet.getRange(DATA_START_ROW, 15, numRows, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function restoreGridShadingForDataRows_(formSheet) { restoreDataRowDesign_(formSheet, false); }
function grayOutDataRowsOnly_(formSheet)           { restoreDataRowDesign_(formSheet, true);  }


// ============== REFRESH J3 DROPDOWN VALIDATION ==============
function refreshDropdownValidation_() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
    const logSheet  = ss.getSheetByName(LOG_SHEET_NAME);
    if (!formSheet || !logSheet) return;

    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;

    const raw     = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const seen    = new Set();
    const records = [];
    for (const row of raw) {
      const key    = String(row[0] || '').trim();
      const status = String(row[2] || '').trim().toUpperCase();
      if (key && /^QC-\d{8}-\d{3}$/.test(key) && status === 'OPEN' && !seen.has(key)) {
        seen.add(key);
        records.push(key);
      }
    }
    if (records.length === 0) return;

    records.sort((a, b) => b.localeCompare(a));

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(records, true)
      .setAllowInvalid(false)
      .setHelpText('Select an OPEN QC Record # to load')
      .build();

    formSheet.getRange(DROPDOWN_CELL_A1).setDataValidation(rule);

  } catch (err) {
    Logger.log('refreshDropdownValidation_ error: ' + err);
  }
}


// ============== GRID CLEARING (DATA ONLY) ==============
function clearGridDataOnly_(formSheet) {
  const lastDataRow = formSheet.getLastRow();
  if (lastDataRow < DATA_START_ROW) return;

  const numRows  = lastDataRow - DATA_START_ROW + 1;
  const width    = END_COL - START_COL + 1;
  const range    = formSheet.getRange(DATA_START_ROW, START_COL, numRows, width);
  const formulas = range.getFormulas();

  range.clearContent();

  const blankValues = range.getValues();
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[0].length; c++) {
      if (formulas[r][c]) blankValues[r][c] = formulas[r][c];
    }
  }
  range.setValues(blankValues);
}


// ============== HELPERS ==============
function findHeaderIndex_(sheet, headerRow, startCol, width, regex) {
  const headers = sheet.getRange(headerRow, startCol, 1, width).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (h && regex.test(h)) return i;
  }
  return -1;
}

function findLastDataRowByToolCode_(sheet, startRow, toolCodeAbsCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;
  const values = sheet.getRange(startRow, toolCodeAbsCol, lastRow - startRow + 1, 1)
    .getDisplayValues().flat();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i] || '').trim() !== '') return startRow + i;
  }
  return startRow - 1;
}

function findExistingRecordRows_(logSheet, recordKey) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return { found: false, rows: [], firstRow: null };
  const keys = logSheet.getRange(2, LOG_COL_RECORDKEY, lastRow - 1, 1).getValues().flat();
  const rows = [];
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i] || '').trim() === String(recordKey).trim()) rows.push(i + 2);
  }
  return { found: rows.length > 0, rows, firstRow: rows.length ? rows[0] : null };
}

function getAllLogRowsForRecord_(logSheet, recordKey) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = logSheet.getLastColumn();
  const data    = logSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out     = [];
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][LOG_COL_RECORDKEY - 1] || '').trim();
    if (key === recordKey) out.push(data[i]);
  }
  out.sort((a, b) => Number(a[LOG_COL_LINEITEM - 1]) - Number(b[LOG_COL_LINEITEM - 1]));
  return out;
}

function makeDailyRecordKey_(logSheet, dateOfMfgDisplay, ss) {
  const tz      = ss.getSpreadsheetTimeZone();
  let baseDate  = new Date();
  const parsed  = new Date(dateOfMfgDisplay);
  if (!isNaN(parsed.getTime())) baseDate = parsed;
  const dateStr = Utilities.formatDate(baseDate, tz, "yyyyMMdd");
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return `QC-${dateStr}-001`;
  const keys    = logSheet.getRange(2, LOG_COL_RECORDKEY, lastRow - 1, 1).getValues().flat();
  let maxSeq    = 0;
  keys.forEach(k => {
    if (!k) return;
    const m = String(k).match(/^QC-(\d{8})-(\d{3})$/);
    if (!m) return;
    if (m[1] === dateStr) {
      const seq = Number(m[2]);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return `QC-${dateStr}-${String(maxSeq + 1).padStart(3, '0')}`;
}

function flashGrid_(ms) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheet      = ss.getSheetByName(FORM_SHEET_NAME);
  const width      = END_COL - START_COL + 1;
  const toolRelIdx = findHeaderIndex_(sheet, HEADER_ROW, START_COL, width, /tool\s*code/i);
  if (toolRelIdx === -1) return;
  const toolAbsCol  = START_COL + toolRelIdx;
  const lastDataRow = findLastDataRowByToolCode_(sheet, DATA_START_ROW, toolAbsCol);
  const flashRange  = sheet.getRange(
    HEADER_ROW, START_COL,
    Math.max(2, lastDataRow - HEADER_ROW + 1), width
  );
  const originalBackgrounds = flashRange.getBackgrounds();
  flashRange.setBackground('#C8E6C9');
  Utilities.sleep(ms || 600);
  flashRange.setBackgrounds(originalBackgrounds);
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  MASTER REGISTER — TOOL CODE LOADER                                     ║
// ║  Reads All Molds tab col C → writes to Data Validation Lists!A3         ║
// ║  F6 dropdown picks it up automatically (no validation rule changes)     ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function getMasterRegId_DF() {
  let id = PropertiesService.getScriptProperties()
             .getProperty("MASTER_REG_ID_DF");
  if (id) return id;
  const st = SpreadsheetApp.getActive()
               .getSheetByName(DF_SETTINGS_SHEET);
  if (st) {
    id = String(st.getRange(DF_MR_ID_CELL).getValue() || "").trim();
    if (id) {
      PropertiesService.getScriptProperties()
        .setProperty("MASTER_REG_ID_DF", id);
      return id;
    }
  }
  throw new Error(
    "Master Register ID not set. " +
    "Run QC Tools → Set Master Register ID and paste the URL."
  );
}

function loadToolCodesIntoList() {
  const ss  = SpreadsheetApp.getActive();
  const dvl = ss.getSheetByName("Data Validation Lists");
  if (!dvl) {
    SpreadsheetApp.getUi().alert(
      "'Data Validation Lists' sheet not found.\n" +
      "Check the tab name matches exactly."
    );
    return;
  }

  let mr, tab;
  try {
    mr  = SpreadsheetApp.openById(getMasterRegId_DF());
    tab = mr.getSheetByName("All Molds");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Could not open Master Register:\n" + e.message);
    return;
  }

  if (!tab) {
    SpreadsheetApp.getUi().alert(
      "'All Molds' tab not found in Master Register.\n" +
      "Check the tab name matches exactly."
    );
    return;
  }

  const lastRow = tab.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data found in the All Molds tab (empty after row 1).");
    return;
  }

  // Col C = Mold Identification (Tool Code)
  const data  = tab.getRange(2, 3, lastRow - 1, 1).getDisplayValues().flat();
  const codes = [];
  for (const val of data) {
    const code = String(val || "").trim();
    if (code && codes.indexOf(code) < 0) codes.push(code);
  }
  codes.sort();

  if (codes.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No tool codes found in column C of the All Molds tab.\n" +
      "Make sure Mold Identification values start from row 2."
    );
    return;
  }

  // Clear old list then write new one starting at A3
  dvl.getRange(3, 1, Math.max(codes.length + 20, 60), 1).clearContent();
  dvl.getRange(3, 1, codes.length, 1).setValues(codes.map(c => [c]));

  SpreadsheetApp.getActive().toast(
    codes.length + " tool codes written to Data Validation Lists!A3:A" +
    (codes.length + 2) + ".  Your F6 dropdown is now updated.",
    "✅ Tool Codes Loaded",
    6
  );
}

function setMasterRegId_DF() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Link Master Register",
    "Paste the full URL or Spreadsheet ID of the Master Register:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  let input = resp.getResponseText().trim();
  const m   = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) input = m[1];
  try {
    SpreadsheetApp.openById(input);
    PropertiesService.getScriptProperties()
      .setProperty("MASTER_REG_ID_DF", input);
    // Also save to Settings tab if it exists
    const st = SpreadsheetApp.getActive().getSheetByName(DF_SETTINGS_SHEET);
    if (st) st.getRange(DF_MR_ID_CELL).setValue(input);
    ui.alert(
      "✅ Master Register linked!\n\n" +
      "Now run QC Tools → Load Tool Codes from Master Register."
    );
  } catch (e) {
    ui.alert("❌ Could not open that spreadsheet:\n" + e.message);
  }
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  ONE-TIME DATA FIX — TOOL CODE NAME STANDARDISATION                    ║
// ║  Run once to align Test Data col J with Master Register names.          ║
// ║  Safe to delete after running.                                          ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function fixTubNames() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const db  = ss.getSheetByName("Test Data");
  if (!db) { SpreadsheetApp.getUi().alert("Test Data sheet not found."); return; }

  const lastRow = db.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data found."); return; }

  // Column J = Tool Code = column 10
  const range  = db.getRange(2, 10, lastRow - 1, 1);
  const values = range.getValues();

  const fixMap = {
    "6.5 Tub (Cav 9-12)"          : "6.5 Tub Cav 9-12",
    "6.5 Tub (Cav 13,14,15, & 17)": "6.5 Tub Cav 13-17",
    "6.5 Tub (Cav 1-4)"           : "6.5 Tub Cav 1-4",
    "48-D"                         : "48D",
    "48-E"                         : "48E",
    "S-45"                         : "S4.5",
  };

  let fixCount = 0;
  const summary = {};

  for (let i = 0; i < values.length; i++) {
    const current = String(values[i][0] || "").trim();
    if (fixMap[current]) {
      summary[current] = (summary[current] || 0) + 1;
      values[i][0] = fixMap[current];
      fixCount++;
    }
  }

  if (fixCount === 0) {
    SpreadsheetApp.getUi().alert(
      "✅ No mismatches found — all names already correct.\n\n" +
      "Safe to delete fixTubNames()."
    );
    return;
  }

  range.setValues(values);

  // Build a readable summary of what changed
  const lines = Object.entries(summary)
    .map(([old, count]) => `  "${old}" → "${fixMap[old]}"  (${count} rows)`)
    .join("\n");

  SpreadsheetApp.getUi().alert(
    "✅ Done!  " + fixCount + " cells updated in Test Data column J.\n\n" +
    "Changes made:\n" + lines + "\n\n" +
    "You can now delete the fixTubNames() function."
  );
}


// ============== QUARTERLY BACKUP ==============
function quarterlyBackup() {
  const sourceFileName = "Plastic Container Freeze Drop System";
  const backupFolderId = "1OCzHvL2l7y_tYqr2gRa-XksllvPiwi0L";

  const today = new Date();
  const year  = today.getFullYear();
  const month = today.getMonth() + 1;
  const day   = today.getDate();

  const isQuarterEnd =
    (month === 3  && day === 31) ||
    (month === 6  && day === 30) ||
    (month === 9  && day === 30) ||
    (month === 12 && day === 31);

  if (!isQuarterEnd) { Logger.log("Not a quarter-end day. No backup created."); return; }

  let quarter;
  if (month === 3)  quarter = "Q1";
  if (month === 6)  quarter = "Q2";
  if (month === 9)  quarter = "Q3";
  if (month === 12) quarter = "Q4";

  const quarterLabel = `${quarter}_${year}`;
  const ss           = SpreadsheetApp.getActive();
  const sheets       = ss.getSheets();
  const backupName   = `${sourceFileName} – Full Backup – ${quarterLabel}`;

  const folder     = DriveApp.getFolderById(backupFolderId);
  const backupFile = SpreadsheetApp.create(backupName);
  DriveApp.getRootFolder().removeFile(backupFile);
  folder.addFile(backupFile);

  const backupSS = SpreadsheetApp.openById(backupFile.getId());
  backupSS.deleteSheet(backupSS.getSheets()[0]);
  sheets.forEach(sh => sh.copyTo(backupSS).setName(sh.getName()));

  Logger.log(`Full backup created: ${backupName}`);
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  QC FORM REDESIGN                                                       ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function redesignQCForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FORM_SHEET_NAME);
  if (!sh) { SpreadsheetApp.getUi().alert('QC Form sheet not found.'); return; }

  SpreadsheetApp.getActive().toast('Applying new form design...', '🎨 Redesign', -1);
  sh.setHiddenGridlines(true);

  const NAVY     = '#1F3864';
  const NAVY2    = '#2E4B7B';
  const NAVY3    = '#3A5A8A';
  const GOLD     = '#C9A027';
  const GOLD_LT  = '#F0E0A0';
  const WHITE    = '#FFFFFF';
  const TEAL     = '#17748A';
  const AMBER    = '#ED7D31';
  const AMBER_LT = '#FFF3E0';
  const GREEN    = '#27AE60';
  const GREEN_LT = '#E8F5E9';
  const BORDER   = '#C5CAD0';
  const BLK      = '#212121';

  function s(cell, bg, fc, bold, size, italic, wrap, halign, valign) {
    cell.setBackground(bg || null);
    cell.setFontColor(fc || BLK);
    cell.setFontWeight(bold ? 'bold' : 'normal');
    cell.setFontSize(size || 10);
    cell.setFontStyle(italic ? 'italic' : 'normal');
    cell.setFontFamily('Calibri');
    if (wrap !== undefined) cell.setWrap(wrap);
    cell.setHorizontalAlignment(halign || 'left');
    cell.setVerticalAlignment(valign || 'middle');
  }

  function border(range, color, style) {
    const c  = color || BORDER;
    const st = style || SpreadsheetApp.BorderStyle.SOLID;
    range.setBorder(true, true, true, true, false, false, c, st);
  }

  sh.setRowHeight(1, 8);
  sh.setRowHeight(2, 56);
  sh.setRowHeight(3, 36);
  sh.setRowHeight(4, 36);
  sh.setRowHeight(5, 90);
  for (let r = 6; r <= 65; r++) sh.setRowHeight(r, 28);

  sh.setColumnWidth(1,  8);
  sh.setColumnWidth(2,  70);
  sh.setColumnWidth(3,  160);
  sh.setColumnWidth(4,  60);
  sh.setColumnWidth(5,  80);
  sh.setColumnWidth(6,  230);
  sh.setColumnWidth(7,  75);
  sh.setColumnWidth(8,  110);
  sh.setColumnWidth(9,  90);
  sh.setColumnWidth(10, 160);
  sh.setColumnWidth(11, 160);
  sh.setColumnWidth(12, 75);
  sh.setColumnWidth(13, 85);
  sh.setColumnWidth(14, 170);
  sh.setColumnWidth(15, 160);
  sh.setColumnWidth(16, 220);
  sh.setColumnWidth(17, 180);
  sh.setColumnWidth(18, 8);

  sh.getRange('A1:R1').setBackground(NAVY);

  sh.getRange('A2').setBackground(NAVY);
  const title = sh.getRange('B2:R2');
  title.merge();
  s(title, NAVY, GOLD, true, 22, false, false, 'center', 'middle');
  title.setValue('🧊  DROP FREEZE TESTING — QC DATA ENTRY FORM');

  sh.getRange('A3:R3').setBackground(NAVY2);

  const recLabel = sh.getRange('B3:C3');
  recLabel.merge();
  s(recLabel, NAVY2, GOLD_LT, true, 12, false, false, 'right', 'middle');
  recLabel.setValue('QC RECORD #');

  const recInput = sh.getRange('D3:F3');
  s(recInput, WHITE, NAVY, true, 12, false, false, 'center', 'middle');
  border(recInput, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange('G3').setBackground(NAVY2);

  const loadLabel = sh.getRange('H3:I3');
  loadLabel.merge();
  s(loadLabel, NAVY2, GOLD_LT, true, 12, false, false, 'right', 'middle');
  loadLabel.setValue('LOAD RECORD');

  const loadInput = sh.getRange('J3:L3');
  s(loadInput, WHITE, TEAL, true, 12, false, false, 'center', 'middle');
  border(loadInput, TEAL, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange('M3:N3').setBackground(NAVY2);

  const saveLabel = sh.getRange('O3');
  s(saveLabel, GREEN, WHITE, true, 12, false, false, 'center', 'middle');
  saveLabel.setValue('✅  SAVE RECORD');
  border(sh.getRange('O3'), GREEN);

  const saveChk = sh.getRange('P3');
  s(saveChk, GREEN_LT, GREEN, true, 12, false, false, 'center', 'middle');
  border(saveChk, GREEN);

  sh.getRange('Q3:R3').setBackground(NAVY2);

  sh.getRange('A4:R4').setBackground(NAVY3);

  const infoLabel = sh.getRange('B4:G4');
  infoLabel.merge();
  s(infoLabel, NAVY3, GOLD_LT, false, 9, true, false, 'left', 'middle');
  infoLabel.setValue('  ℹ️  Fill in sample rows below, then check SAVE RECORD. To update an existing record, select it from LOAD RECORD first.');

  sh.getRange('H4:N4').setBackground(NAVY3);

  const clearLabel = sh.getRange('O4');
  s(clearLabel, AMBER, WHITE, true, 12, false, false, 'center', 'middle');
  clearLabel.setValue('🗑️  CLEAR FORM');
  border(sh.getRange('O4'), AMBER);

  const clearChk = sh.getRange('P4');
  s(clearChk, AMBER_LT, AMBER, true, 12, false, false, 'center', 'middle');
  border(clearChk, AMBER);

  sh.getRange('Q4:R4').setBackground(NAVY3);

  sh.getRange('A5').setBackground(NAVY);
  sh.getRange('R5').setBackground(NAVY);

  [
    { col: 2,  label: 'Line #' },
    { col: 3,  label: 'Date / Time of Mfg\n(01/01/1999 6:00 am)' },
    { col: 4,  label: 'Shift' },
    { col: 5,  label: 'Pallet #' },
    { col: 6,  label: 'Tool Code' },
    { col: 7,  label: 'Cavity' },
    { col: 8,  label: 'Resin ID' },
    { col: 9,  label: 'Test Type' },
  ].forEach(h => {
    const cell = sh.getRange(5, h.col);
    cell.setValue(h.label);
    s(cell, AMBER, WHITE, true, 12, false, true, 'center', 'middle');
    cell.setBorder(true, true, true, true, false, false, '#B35E00', SpreadsheetApp.BorderStyle.SOLID);
  });

  [
    { col: 10, label: 'Test Date / Time\n(01/01/1999 6:00 am)' },
    { col: 11, label: 'Tested By' },
    { col: 12, label: 'Freezer\nTemp (°F)' },
    { col: 13, label: 'Drop\nHeight (in)' },
    { col: 14, label: 'Drop Angle' },
    { col: 15, label: 'Pass / Fail\nResult' },
    { col: 16, label: 'Failure Description\n(wall side & observations)' },
    { col: 17, label: 'Other Notes /\nObservations' },
  ].forEach(h => {
    const cell = sh.getRange(5, h.col);
    cell.setValue(h.label);
    s(cell, TEAL, WHITE, true, 12, false, true, 'center', 'middle');
    cell.setBorder(true, true, true, true, false, false, '#0E5060', SpreadsheetApp.BorderStyle.SOLID);
  });

  restoreDataRowDesign_(sh, false);

  sh.getRange('A66:R66').setBackground(NAVY);
  sh.setFrozenRows(5);

  SpreadsheetApp.getActive().toast('✅ QC Form redesign complete!', '🎨 Done', 5);
  ss.setActiveSheet(sh);
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  MENU                                                                   ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QC Tools')
    .addItem('🎨 Redesign QC Form',                    'redesignQCForm')
    .addSeparator()
    .addItem('🔨 Build / Rebuild All Dashboards',       'buildDashboard')
    .addSeparator()
    .addItem('🔄 Load Tool Codes from Master Register', 'loadToolCodesIntoList')
    .addItem('🔗 Set Master Register ID',               'setMasterRegId_DF')
    .addToUi();
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  FREEZE TEST QUALITY DASHBOARD SYSTEM (v9)                              ║
// ╚══════════════════════════════════════════════════════════════════════════╝

const DASH_SHEET  = '📊 DASHBOARD';
const AGING_SHEET = '⏳ In Progress';
const MOLD_SHEET  = '🔍 MOLD ANALYSIS';
const SRC_SHEET   = 'Test Data';

const TD = {
  RECORD:   'A',
  LINEITEM: 'B',
  STATUS:   'C',
  CREATED:  'D',
  UPDATED:  'E',
  LINENUM:  'F',
  DATEMFG:  'G',
  SHIFT:    'H',
  PALLET:   'I',
  TOOLCODE: 'J',
  CAVITY:   'K',
  RESIN:    'L',
  TESTTYPE: 'M',
  TESTDATE: 'N',
  TESTEDBY: 'O',
  TEMP:     'P',
  HEIGHT:   'Q',
  ANGLE:    'R',
  RESULT:   'S',
  FAILDESC: 'T',
  NOTES:    'U',
  YEAR:     'V',
  MONTH:    'W',
};

const CLR = {
  NAVY:      '#1F3864',
  NAVY2:     '#2E4B7B',
  GOLD:      '#C9A027',
  GOLD_LT:   '#F0E0A0',
  WHITE:     '#FFFFFF',
  LTGRAY:    '#F5F5F5',
  GREEN_DK:  '#375623',
  GREEN_BG:  '#E2EFDA',
  RED_DK:    '#9C0006',
  RED_BG:    '#FFE2E2',
  AMBER:     '#ED7D31',
  AMBER_BG:  '#FFF2CC',
  TEAL:      '#17748A',
  TEAL_BG:   '#D6EAF8',
  PURPLE:    '#6B3FA0',
  PURPLE_BG: '#EDE7F6',
  BLK:       '#000000',
  GRAY2:     '#AAAAAA',
};

const ALL_MOLDS = [
  '6.5 Tub Cav 1-4',
  '6.5 Tub Cav 9-12',
  '6.5 Tub Cav 13-17',
  '3 Gal I/C Tub',
  '48A',
  '48D',
  '49',
  'S3',
];
const TUB_MOLDS = [
  '6.5 Tub Cav 1-4',
  '6.5 Tub Cav 9-12',
  '6.5 Tub Cav 13-17',
];
const ALL_ANGLES = [
  'Flat- Bottom',
  '45 deg- Bottom',
  'Flat- Top',
  '45 deg- Top',
  'Side Drop',
];
const TEST_TYPES = [
  { label: '24 hr (Blast Freezer)', pattern: '24*' },
  { label: '72 hr (Standard Freezer)', pattern: '72*' },
  { label: '8 hr (Rapid)', pattern: '8*' },
];
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];


function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActive().toast('Building dashboard system...', '📊 Dashboard', -1);
  buildSummaryDash_(ss);
  buildAgingSheet_(ss);
  buildMoldAnalysis_(ss);
  setupDashTrigger_();
  ss.setActiveSheet(ss.getSheetByName(DASH_SHEET));
  SpreadsheetApp.getActive().toast(
    '✅ All 3 dashboard tabs ready! Formulas update live.',
    'Done', 6
  );
}

function refreshDashboard() {
  SpreadsheetApp.getActive().toast('✅ All dashboards update live automatically — no refresh needed.', 'Live Data', 5);
}


function buildSummaryDash_(ss) {
  const sh  = resetSheet_(ss, DASH_SHEET, CLR.NAVY.replace('#', ''), 0);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(3);

  [20, 185, 175, 108, 108, 108, 108, 108, 108, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const SRC = SRC_SHEET;
  const AG  = AGING_SHEET;
  let row = 1;

  rh_(sh, row, 10); fillRow_(sh, row, 2, 9, CLR.NAVY); row++;
  rh_(sh, row, 62);
  mw_(sh, row, 2, 9)
    .setValue('🧊  FREEZE DROP TEST — QUALITY DASHBOARD')
    .setFontFamily('Calibri').setFontSize(26).setFontWeight('bold')
    .setFontColor(CLR.GOLD).setBackground(CLR.NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 22);
  mw_(sh, row, 2, 9)
    .setFormula(`="Data refreshed: "&TEXT(NOW(),"mmm dd, yyyy  h:mm AM/PM")`)
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
    .setFontColor(CLR.GOLD_LT).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 10); fillRow_(sh, row, 2, 9, CLR.NAVY); row++;

  rh_(sh, row, 26);
  mw_(sh, row, 2, 9)
    .setValue('KEY PERFORMANCE INDICATORS')
    .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  const totF = `COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`
             + `+COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`
             + `+COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
  const pasF = `COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`;
  const faiF = `COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`;
  const incF = `COUNTIFS('${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;

  const opF  = `COUNTA('${AG}'!J14:J)`;
  const pdF  = `COUNTIF('${AG}'!J14:J,"🔴  PAST DUE")`;

  const kpiSets = [
    [
      { label:'TOTAL TESTS COMPLETED',   valF:`=${totF}`, subF:`=${pasF}&" Passed   |   "&${faiF}&" Failed   |   "&${incF}&" Inconclusive"`, vColor:CLR.GOLD,  hColor:CLR.NAVY2 },
      { label:'OVERALL PASS RATE',       valF:`=IFERROR(TEXT((${pasF})/(${totF}),"0.0%"),"N/A")`, subF:`="Pass rate excludes Inconclusive results"`, vColor:'#00C851', hColor:CLR.NAVY2 },
    ],
    [
      { label:'OPEN — AWAITING RESULTS', valF:`=${opF}`,  subF:`="Samples in freezer or pending test"`,  vColor:CLR.AMBER, hColor:CLR.NAVY2  },
      { label:'⚠️  PAST DUE',           valF:`=${pdF}`,  subF:`="Tests overdue — see In Progress tab"`, vColor:'#FF4444', hColor:'#9C0006'  },
    ],
  ];

  for (const kpiRow of kpiSets) {
    rh_(sh, row, 10); fillRow_(sh, row, 2, 9, CLR.NAVY); row++;
    rh_(sh, row, 28); rh_(sh, row+1, 54); rh_(sh, row+2, 18);
    kpiRow.forEach((k, idx) => {
      const sc = idx === 0 ? 2 : 6;
      const ec = idx === 0 ? 4 : 9;
      mw_(sh, row,   sc, ec).setValue(k.label).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.WHITE).setBackground(k.hColor).setHorizontalAlignment('center').setVerticalAlignment('middle');
      mw_(sh, row+1, sc, ec).setFormula(k.valF).setFontFamily('Calibri').setFontSize(34).setFontWeight('bold').setFontColor(k.vColor).setBackground(CLR.NAVY).setHorizontalAlignment('center').setVerticalAlignment('middle');
      mw_(sh, row+2, sc, ec).setFormula(k.subF).setFontFamily('Calibri').setFontSize(9).setFontStyle('italic').setFontColor(CLR.GRAY2).setBackground(CLR.NAVY).setHorizontalAlignment('center').setVerticalAlignment('middle');
    });
    for (let r = row; r <= row+2; r++) sh.getRange(r, 5).setBackground(CLR.NAVY);
    row += 3;
  }

  rh_(sh, row, 10); fillRow_(sh, row, 2, 9, CLR.NAVY); row++;
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 9, '📦   PASS / FAIL RATES BY MOLD');
  row = tableHdrs_(sh, row, [{sc:2,ec:3,l:'Mold / Tool Code'},{sc:4,ec:4,l:'Total Tests'},{sc:5,ec:5,l:'Passed'},{sc:6,ec:6,l:'Failed'},{sc:7,ec:7,l:'Inconclusive'},{sc:8,ec:8,l:'Pass Rate'},{sc:9,ec:9,l:'Fail Rate'}]);
  ALL_MOLDS.forEach((mold, i) => {
    const bg = i % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
    const mQ = mold.replace(/"/g, '""');
    const base = `'${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.TOOLCODE}:${TD.TOOLCODE},"${mQ}"`;
    const t = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    const p = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`;
    const f = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`;
    const ic= `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    rh_(sh, row, 22);
    mw_(sh, row, 2, 3).setValue(mold).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
    cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
    cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
    cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
    cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
    cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
    cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
    borders_(sh, row, 2, 9);
    row++;
  });
  row = grandTotal_(sh, row, 2, 9, 'Grand Total — All Molds', totF, pasF, faiF, incF);
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 9, '📐   PASS / FAIL RATES BY DROP ANGLE');
  row = tableHdrs_(sh, row, [{sc:2,ec:3,l:'Drop Angle'},{sc:4,ec:4,l:'Total Tests'},{sc:5,ec:5,l:'Passed'},{sc:6,ec:6,l:'Failed'},{sc:7,ec:7,l:'Inconclusive'},{sc:8,ec:8,l:'Pass Rate'},{sc:9,ec:9,l:'Fail Rate'}]);
  ALL_ANGLES.forEach((angle, i) => {
    const bg = i % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
    const base = `'${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.ANGLE}:${TD.ANGLE},"${angle}"`;
    const t = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    const p = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`;
    const f = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`;
    const ic= `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    rh_(sh, row, 22);
    mw_(sh, row, 2, 3).setValue(angle).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
    cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
    cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
    cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
    cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
    cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
    cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
    borders_(sh, row, 2, 9);
    row++;
  });
  row = grandTotal_(sh, row, 2, 9, 'Grand Total — All Angles', totF, pasF, faiF, incF);
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 9, '🧊   6.5 TUB ANALYSIS — DROP ANGLE × MOLD CAVITY GROUP');
  row = tableHdrs_(sh, row, [{sc:2,ec:2,l:'Drop Angle'},{sc:3,ec:3,l:'Mold / Cavity Group'},{sc:4,ec:4,l:'Total'},{sc:5,ec:5,l:'Passed'},{sc:6,ec:6,l:'Failed'},{sc:7,ec:7,l:'Inconclusive'},{sc:8,ec:8,l:'Pass Rate'},{sc:9,ec:9,l:'Fail Rate'}]);
  let altIdx = 0;
  for (const angle of ALL_ANGLES) {
    for (const mold of TUB_MOLDS) {
      const bg = altIdx % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
      const mQ = mold.replace(/"/g, '""');
      const base = `'${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.ANGLE}:${TD.ANGLE},"${angle}",'${SRC}'!${TD.TOOLCODE}:${TD.TOOLCODE},"${mQ}"`;
      const t = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
      const p = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`;
      const f = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`;
      const ic= `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
      rh_(sh, row, 22);
      sh.getRange(row, 2).setValue(angle).setFontFamily('Calibri').setFontSize(10).setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
      sh.getRange(row, 3).setValue(mold).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.NAVY2).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
      cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
      cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
      cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
      cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
      cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
      cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
      borders_(sh, row, 2, 9);
      row++; altIdx++;
    }
  }
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 9, '📅   PASS / FAIL RATES BY YEAR');
  row = tableHdrs_(sh, row, [{sc:2,ec:3,l:'Year'},{sc:4,ec:4,l:'Total Tests'},{sc:5,ec:5,l:'Passed'},{sc:6,ec:6,l:'Failed'},{sc:7,ec:7,l:'Inconclusive'},{sc:8,ec:8,l:'Pass Rate'},{sc:9,ec:9,l:'Fail Rate'}]);
  const currentYear = new Date().getFullYear();
  for (let yr = 2025; yr <= currentYear; yr++) {
    const i  = yr - 2025;
    const bg = i % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
    const base = `'${SRC}'!${TD.STATUS}:${TD.STATUS},"COMPLETE",'${SRC}'!${TD.YEAR}:${TD.YEAR},${yr}`;
    const t = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")+COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    const p = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Pass")`;
    const f = `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Fail")`;
    const ic= `COUNTIFS(${base},'${SRC}'!${TD.RESULT}:${TD.RESULT},"Inconclusive")`;
    rh_(sh, row, 22);
    mw_(sh, row, 2, 3).setValue(yr).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
    cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
    cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
    cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
    cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
    cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
    cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
    borders_(sh, row, 2, 9);
    row++;
  }
  row = grandTotal_(sh, row, 2, 9, 'Grand Total — All Years', totF, pasF, faiF, incF);
  rh_(sh, row, 18); row++;

  rh_(sh, row, 22);
  mw_(sh, row, 2, 9)
    .setValue('★  Quality Management  |  All formulas update live  |  VOID records excluded from all counts')
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
    .setFontColor(CLR.NAVY2).setBackground(CLR.LTGRAY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}


function buildAgingSheet_(ss) {
  const sh  = resetSheet_(ss, AGING_SHEET, 'ED7D31', 1);
  sh.setHiddenGridlines(true);

  [20, 155, 210, 65, 110, 90, 155, 155, 115, 150, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const OV = 'Open View';
  let row = 1;

  rh_(sh, row, 10); fillRow_(sh, row, 2, 10, CLR.NAVY); row++;
  rh_(sh, row, 58);
  mw_(sh, row, 2, 10)
    .setValue('⏳  IN PROGRESS TESTS — AGING & PAST DUE TRACKER')
    .setFontFamily('Calibri').setFontSize(22).setFontWeight('bold')
    .setFontColor(CLR.GOLD).setBackground(CLR.NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 22);
  mw_(sh, row, 2, 10)
    .setFormula(`="🟢 Live as of: "&TEXT(NOW(),"mmm dd, yyyy  h:mm AM/PM")&"   |   Source: Open View tab"`)
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
    .setFontColor(CLR.GOLD_LT).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 10); fillRow_(sh, row, 2, 10, CLR.NAVY); row++;

  rh_(sh, row, 26);
  mw_(sh, row, 2, 10)
    .setValue('LIVE STATUS SUMMARY')
    .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  const DR = 14;

  const kCards = [
    { label:'OPEN TESTS',    valF:`=COUNTA(J${DR}:J)`,                          color:CLR.AMBER,  bg:CLR.NAVY2 },
    { label:'ON TRACK',      valF:`=COUNTIF(J${DR}:J,"🟢  ON TRACK")`,          color:'#4FC3F7',  bg:CLR.NAVY2 },
    { label:'⚠️  PAST DUE', valF:`=COUNTIF(J${DR}:J,"🔴  PAST DUE")`,          color:'#FF4444',  bg:'#9C0006' },
  ];
  const kCols = [[2,3],[5,7],[8,10]];

  rh_(sh, row, 26);
  kCards.forEach((k, i) => {
    const [sc, ec] = kCols[i];
    mw_(sh, row, sc, ec).setValue(k.label)
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(CLR.WHITE).setBackground(k.bg)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.getRange(row, 4).setBackground(CLR.NAVY);
  row++;

  rh_(sh, row, 54);
  kCards.forEach((k, i) => {
    const [sc, ec] = kCols[i];
    mw_(sh, row, sc, ec).setFormula(k.valF)
      .setFontFamily('Calibri').setFontSize(32).setFontWeight('bold')
      .setFontColor(k.color).setBackground(CLR.NAVY)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.getRange(row, 4).setBackground(CLR.NAVY);
  row++;

  rh_(sh, row, 10); fillRow_(sh, row, 2, 10, CLR.NAVY); row++;

  rh_(sh, row, 22);
  mw_(sh, row, 2, 10).setValue('COLOR LEGEND')
    .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  [
    {bg:CLR.RED_BG, color:CLR.RED_DK, txt:'🔴  PAST DUE — Test window has expired. Perform the test and record results immediately.'},
    {bg:CLR.TEAL_BG, color:CLR.TEAL,  txt:'🟢  ON TRACK — Sample is within its freeze window. Test when due date is reached.'},
  ].forEach(leg => {
    rh_(sh, row, 22);
    mw_(sh, row, 2, 10).setValue(leg.txt)
      .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
      .setFontColor(leg.color).setBackground(leg.bg)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    row++;
  });

  rh_(sh, row, 10); fillRow_(sh, row, 2, 10, CLR.NAVY); row++;

  const hdrRow = row;
  rh_(sh, row, 28);
  [
    {col:2,  label:'QC Record #'},
    {col:3,  label:'Tool Code / Mold'},
    {col:4,  label:'Cav'},
    {col:5,  label:'Date of Mfg'},
    {col:6,  label:'Test Type'},
    {col:7,  label:'Created'},
    {col:8,  label:'Due Date & Time'},
    {col:9,  label:'Hrs Remaining'},
    {col:10, label:'⚠️ AGING STATUS'},
  ].forEach(h => {
    sh.getRange(row, h.col)
      .setValue(h.label)
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(CLR.WHITE)
      .setBackground(h.label.includes('AGING') ? '#9C0006' : CLR.TEAL)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  });
  row++;

  sh.setFrozenRows(hdrRow);

  if (row !== DR) {
    Logger.log('⚠️ buildAgingSheet_: DR constant (' + DR + ') != actual data row (' + row + '). Update DR constant.');
  }

  const fc = `'${OV}'!A2:A<>""`;

  sh.getRange(DR, 2).setFormula(`=IFERROR(FILTER('${OV}'!A2:A,${fc}),"✅  No open tests — all clear!")`);
  sh.getRange(DR, 3).setFormula(`=IFERROR(FILTER('${OV}'!J2:J,${fc}),"")`);
  sh.getRange(DR, 4).setFormula(`=IFERROR(FILTER('${OV}'!K2:K,${fc}),"")`);
  sh.getRange(DR, 5).setFormula(`=IFERROR(FILTER('${OV}'!G2:G,${fc}),"")`);
  sh.getRange(DR, 5).setNumberFormat('mm/dd/yyyy');
  sh.getRange(DR, 6).setFormula(`=IFERROR(FILTER('${OV}'!M2:M,${fc}),"")`);
  sh.getRange(DR, 7).setFormula(`=IFERROR(FILTER('${OV}'!D2:D,${fc}),"")`);
  sh.getRange(DR, 7).setNumberFormat('mm/dd/yy h:mm am/pm');
  sh.getRange(DR, 8).setFormula(
    `=IFERROR(ARRAYFORMULA(IF(G${DR}:G="","",` +
    `G${DR}:G+IF(ISNUMBER(SEARCH("72",TO_TEXT(F${DR}:F))),3,` +
    `IF(ISNUMBER(SEARCH("24",TO_TEXT(F${DR}:F))),1,8/24)))),"")` );
  sh.getRange(DR, 8).setNumberFormat('mm/dd/yy h:mm am/pm');
  sh.getRange(DR, 9).setFormula(
    `=IFERROR(ARRAYFORMULA(IF(H${DR}:H="","",ROUND((H${DR}:H-NOW())*24,1))),"")` );
  sh.getRange(DR, 9).setNumberFormat('0.0" hrs"');
  sh.getRange(DR, 10).setFormula(
    `=IFERROR(ARRAYFORMULA(IF(H${DR}:H="","",IF(H${DR}:H<NOW(),"🔴  PAST DUE","🟢  ON TRACK"))),"")` );

  sh.getRange(DR, 2, 500, 9)
    .setFontFamily('Calibri').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J${DR}="🔴  PAST DUE"`)
    .setBackground(CLR.RED_BG).setFontColor(CLR.RED_DK)
    .setRanges([sh.getRange(DR, 2, 500, 8)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J${DR}="🟢  ON TRACK"`)
    .setBackground(CLR.TEAL_BG)
    .setRanges([sh.getRange(DR, 2, 500, 8)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J${DR}="🔴  PAST DUE"`)
    .setBackground('#9C0006').setFontColor(CLR.WHITE).setBold(true)
    .setRanges([sh.getRange(DR, 10, 500, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$J${DR}="🟢  ON TRACK"`)
    .setBackground(CLR.GREEN_DK).setFontColor(CLR.WHITE).setBold(true)
    .setRanges([sh.getRange(DR, 10, 500, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(CLR.RED_DK).setBold(true)
    .setRanges([sh.getRange(DR, 9, 500, 1)]).build());
  sh.setConditionalFormatRules(rules);
}


function buildMoldAnalysis_(ss) {
  const sh = resetSheet_(ss, MOLD_SHEET, '17748A', 2);
  sh.setHiddenGridlines(true);

  [20, 165, 145, 95, 95, 95, 95, 95, 95, 105, 85, 85, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const SRC = SRC_SHEET;
  const MC  = '$E$5';
  let row = 1;

  rh_(sh, row, 10); fillRow_(sh, row, 2, 12, CLR.NAVY); row++;
  rh_(sh, row, 58);
  mw_(sh, row, 2, 12)
    .setValue('🔍  MOLD DRILL-DOWN ANALYSIS')
    .setFontFamily('Calibri').setFontSize(24).setFontWeight('bold')
    .setFontColor(CLR.GOLD).setBackground(CLR.NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 22);
  mw_(sh, row, 2, 12)
    .setFormula(`="Select a mold to filter all tables below  |  Data refreshed: "&TEXT(NOW(),"mmm dd, yyyy  h:mm AM/PM")`)
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
    .setFontColor(CLR.GOLD_LT).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;
  rh_(sh, row, 10); fillRow_(sh, row, 2, 12, CLR.NAVY); row++;

  rh_(sh, row, 38);
  mw_(sh, row, 2, 4)
    .setValue('🔎  FILTER BY MOLD:')
    .setFontFamily('Calibri').setFontSize(12).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('right').setVerticalAlignment('middle');

  const dropRange = mw_(sh, row, 5, 8);
  dropRange.setValue('ALL')
    .setFontFamily('Calibri').setFontSize(14).setFontWeight('bold')
    .setFontColor(CLR.NAVY).setBackground(CLR.WHITE)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, CLR.TEAL, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const moldList = ['ALL', ...ALL_MOLDS];
  const moldRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(moldList, true)
    .setAllowInvalid(false)
    .setHelpText('Select a mold or ALL to show everything')
    .build();
  dropRange.setDataValidation(moldRule);

  const moldTot = mfTotal_(SRC, MC, '');
  mw_(sh, row, 9, 12)
    .setFormula(`=IF(${MC}="ALL","Showing all molds  |  "&${moldTot}&" tests","Showing: "&${MC}&"  |  "&${moldTot}&" tests")`)
    .setFontFamily('Calibri').setFontSize(10).setFontStyle('italic')
    .setFontColor(CLR.GOLD_LT).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  rh_(sh, row, 10); fillRow_(sh, row, 2, 12, CLR.NAVY); row++;
  rh_(sh, row, 10); row++;

  row = sectionHdr_(sh, row, 2, 12, '📐   PASS / FAIL BY DROP ANGLE — FILTERED BY MOLD');
  row = tableHdrs12_(sh, row, ['Drop Angle','','Total','Pass','Fail','Inc','Pass %','Fail %','','','','']);
  ALL_ANGLES.forEach((angle, i) => {
    const bg    = i % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
    const extra = `'${SRC}'!${TD.ANGLE}:${TD.ANGLE},"${angle}"`;
    const t = mfTotal_(SRC, MC, extra);
    const p = mfc_(SRC, MC, extra, 'Pass');
    const f = mfc_(SRC, MC, extra, 'Fail');
    const ic= mfc_(SRC, MC, extra, 'Inconclusive');
    rh_(sh, row, 22);
    mw_(sh, row, 2, 3).setValue(angle).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
    cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
    cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
    cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
    cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
    cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
    cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
    mw_(sh, row, 10, 12).setBackground(bg);
    borders_(sh, row, 2, 9);
    row++;
  });
  const angTot = mfTotal_(SRC, MC, '');
  const angP   = mfc_(SRC, MC, '', 'Pass');
  const angF   = mfc_(SRC, MC, '', 'Fail');
  const angI   = mfc_(SRC, MC, '', 'Inconclusive');
  row = grandTotal12_(sh, row, angTot, angP, angF, angI);
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 12, '🌡️   PASS / FAIL BY TEST TYPE + FREEZER TEMPERATURE — FILTERED BY MOLD');
  rh_(sh, row, 26);
  ['Test Type','','Total','Pass','Fail','Inc','Pass %','Fail %','Avg Temp °F','Min Temp','Max Temp',''].forEach((lbl, i) => {
    const c = sh.getRange(row, i + 2);
    if (lbl === '') { c.setBackground(CLR.TEAL); return; }
    c.setValue(lbl)
     .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
     .setFontColor(CLR.WHITE).setBackground(CLR.TEAL)
     .setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  });
  mw_(sh, row, 2, 3).setValue('Test Type')
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.TEAL)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  row++;

  TEST_TYPES.forEach((tt, i) => {
    const bg    = i % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
    const extra = `'${SRC}'!${TD.TESTTYPE}:${TD.TESTTYPE},"${tt.pattern}"`;
    const t  = mfTotal_(SRC, MC, extra);
    const p  = mfc_(SRC, MC, extra, 'Pass');
    const f  = mfc_(SRC, MC, extra, 'Fail');
    const ic = mfc_(SRC, MC, extra, 'Inconclusive');
    rh_(sh, row, 22);
    mw_(sh, row, 2, 3).setValue(tt.label).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
    cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
    cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
    cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
    cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
    cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
    cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
    cf_(sh, row, 10, `=${mfAvg_(SRC, MC, extra)}`, CLR.TEAL_BG, CLR.TEAL, true, '0.0');
    cf_(sh, row, 11, `=${mfMin_(SRC, MC, extra)}`, CLR.TEAL_BG, CLR.TEAL, false, '0.0');
    cf_(sh, row, 12, `=${mfMax_(SRC, MC, extra)}`, CLR.TEAL_BG, CLR.TEAL, false, '0.0');
    borders_(sh, row, 2, 12);
    row++;
  });
  row = grandTotalTemp_(sh, row, SRC, MC);
  rh_(sh, row, 18); row++;

  row = sectionHdr_(sh, row, 2, 12, '📅   MONTHLY PASS / FAIL TREND — FILTERED BY MOLD');
  row = tableHdrs12_(sh, row, ['Year','Month','Total','Pass','Fail','Inc','Pass %','Fail %','','','','']);

  const firstMonthRow = row;
  const years = [2025, 2026];
  let altM = 0;
  for (const yr of years) {
    rh_(sh, row, 24);
    mw_(sh, row, 2, 12).setValue(`── ${yr} ──`)
      .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
      .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    row++;
    for (let m = 1; m <= 12; m++) {
      const bg    = altM % 2 === 0 ? CLR.LTGRAY : CLR.WHITE;
      const extra = `'${SRC}'!${TD.YEAR}:${TD.YEAR},${yr},'${SRC}'!${TD.MONTH}:${TD.MONTH},${m}`;
      const t  = mfTotal_(SRC, MC, extra);
      const p  = mfc_(SRC, MC, extra, 'Pass');
      const f  = mfc_(SRC, MC, extra, 'Fail');
      const ic = mfc_(SRC, MC, extra, 'Inconclusive');
      rh_(sh, row, 20);
      sh.getRange(row, 2).setValue(yr).setFontFamily('Calibri').setFontSize(9).setFontColor(CLR.GRAY2).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
      sh.getRange(row, 3).setValue(MONTH_NAMES[m-1]).setFontFamily('Calibri').setFontSize(10).setFontWeight('bold').setFontColor(CLR.BLK).setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
      cf_(sh, row, 4, `=${t}`, bg, CLR.BLK, false, null);
      cf_(sh, row, 5, `=${p}`, CLR.GREEN_BG, CLR.GREEN_DK, true, null);
      cf_(sh, row, 6, `=${f}`, CLR.RED_BG, CLR.RED_DK, true, null);
      cf_(sh, row, 7, `=${ic}`, CLR.PURPLE_BG, CLR.PURPLE, true, null);
      cf_(sh, row, 8, `=IFERROR((${p})/(${t}),"—")`, CLR.GREEN_BG, CLR.GREEN_DK, true, '0.0%');
      cf_(sh, row, 9, `=IFERROR((${f})/(${t}),"—")`, CLR.RED_BG, CLR.RED_DK, true, '0.0%');
      mw_(sh, row, 10, 12).setBackground(bg);
      borders_(sh, row, 2, 9);
      row++; altM++;
    }
  }
  const lastMonthRow = row - 1;
  rh_(sh, row, 14); row++;

  rh_(sh, row, 28);
  mw_(sh, row, 2, 12).setValue('📈  PASS RATE TREND  &  TEST VOLUME TREND')
    .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  rh_(sh, row, 50);
  sh.getRange(row, 2).setValue('Pass Rate →')
    .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
    .setFontColor(CLR.GREEN_DK).setBackground(CLR.GREEN_BG)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  mw_(sh, row, 3, 12)
    .setFormula(`=IFERROR(SPARKLINE(FILTER(H${firstMonthRow}:H${lastMonthRow},D${firstMonthRow}:D${lastMonthRow}>0),{"charttype","line";"color","#375623";"linewidth",2}),"No data yet")`)
    .setBackground(CLR.GREEN_BG).setVerticalAlignment('middle');
  row++;

  rh_(sh, row, 50);
  sh.getRange(row, 2).setValue('Volume →')
    .setFontFamily('Calibri').setFontSize(9).setFontWeight('bold')
    .setFontColor(CLR.NAVY).setBackground(CLR.TEAL_BG)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  mw_(sh, row, 3, 12)
    .setFormula(`=IFERROR(SPARKLINE(FILTER(D${firstMonthRow}:D${lastMonthRow},D${firstMonthRow}:D${lastMonthRow}>0),{"charttype","bar";"color1","#1F3864"}),"No data yet")`)
    .setBackground(CLR.TEAL_BG).setVerticalAlignment('middle');
  row++;

  rh_(sh, row, 14); row++;

  rh_(sh, row, 22);
  mw_(sh, row, 2, 12)
    .setValue('★  All tables filter by the mold selected above  |  Formulas update live  |  VOID records excluded  |  Test type wildcards: "24*" matches both "24 hrs" and "24"')
    .setFontFamily('Calibri').setFontSize(9).setFontStyle('italic')
    .setFontColor(CLR.NAVY2).setBackground(CLR.LTGRAY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setFrozenRows(5);
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  MOLD ANALYSIS FORMULA HELPERS                                          ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function mfc_(src, mc, extra, result) {
  const b = `'${src}'!${TD.STATUS}:${TD.STATUS},"COMPLETE"` + (extra ? ',' + extra : '') + `,'${src}'!${TD.RESULT}:${TD.RESULT},"${result}"`;
  const m = `,'${src}'!${TD.TOOLCODE}:${TD.TOOLCODE},${mc}`;
  return `IF(${mc}="ALL",COUNTIFS(${b}),COUNTIFS(${b}${m}))`;
}

function mfTotal_(src, mc, extra) {
  return `(${mfc_(src,mc,extra,'Pass')}+${mfc_(src,mc,extra,'Fail')}+${mfc_(src,mc,extra,'Inconclusive')})`;
}

function mfAvg_(src, mc, extra) {
  const b = `'${src}'!${TD.TEMP}:${TD.TEMP},'${src}'!${TD.STATUS}:${TD.STATUS},"COMPLETE"` + (extra ? ',' + extra : '') + `,'${src}'!${TD.TEMP}:${TD.TEMP},"<>"`;
  const m = `,'${src}'!${TD.TOOLCODE}:${TD.TOOLCODE},${mc}`;
  return `IFERROR(IF(${mc}="ALL",AVERAGEIFS(${b}),AVERAGEIFS(${b}${m})),"—")`;
}

function mfMin_(src, mc, extra) {
  const b = `'${src}'!${TD.TEMP}:${TD.TEMP},'${src}'!${TD.STATUS}:${TD.STATUS},"COMPLETE"` + (extra ? ',' + extra : '') + `,'${src}'!${TD.TEMP}:${TD.TEMP},"<>"`;
  const m = `,'${src}'!${TD.TOOLCODE}:${TD.TOOLCODE},${mc}`;
  return `IFERROR(IF(${mc}="ALL",MINIFS(${b}),MINIFS(${b}${m})),"—")`;
}

function mfMax_(src, mc, extra) {
  const b = `'${src}'!${TD.TEMP}:${TD.TEMP},'${src}'!${TD.STATUS}:${TD.STATUS},"COMPLETE"` + (extra ? ',' + extra : '') + `,'${src}'!${TD.TEMP}:${TD.TEMP},"<>"`;
  const m = `,'${src}'!${TD.TOOLCODE}:${TD.TOOLCODE},${mc}`;
  return `IFERROR(IF(${mc}="ALL",MAXIFS(${b}),MAXIFS(${b}${m})),"—")`;
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  TRIGGER SETUP                                                          ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function setupDashTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshDashboard')
    .forEach(t => ScriptApp.deleteTrigger(t));
}


// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  PRIVATE FORMATTING HELPERS                                             ║
// ╚══════════════════════════════════════════════════════════════════════════╝

function resetSheet_(ss, name, tabColorHex, position) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name, position);
  sh.setTabColor('#' + tabColorHex.replace('#', ''));
  return sh;
}

function mw_(sh, row, startCol, endCol) {
  const r = sh.getRange(row, startCol, 1, endCol - startCol + 1);
  r.merge();
  return r;
}

function rh_(sh, row, height)             { sh.setRowHeight(row, height); }
function fillRow_(sh, row, sc, ec, color) { sh.getRange(row, sc, 1, ec - sc + 1).setBackground(color); }

function cf_(sh, row, col, formula, bg, color, bold, fmt) {
  const c = sh.getRange(row, col);
  c.setFormula(formula)
   .setFontFamily('Calibri').setFontSize(10)
   .setFontWeight(bold ? 'bold' : 'normal')
   .setFontColor(color).setBackground(bg)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (fmt) c.setNumberFormat(fmt);
}

function sectionHdr_(sh, row, sc, ec, title) {
  rh_(sh, row, 30);
  mw_(sh, row, sc, ec)
    .setValue(title)
    .setFontFamily('Calibri').setFontSize(12).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  return row + 1;
}

function tableHdrs_(sh, row, headers) {
  rh_(sh, row, 26);
  headers.forEach(h => {
    const r = (h.sc === h.ec) ? sh.getRange(row, h.sc) : mw_(sh, row, h.sc, h.ec);
    r.setValue(h.l)
     .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
     .setFontColor(CLR.WHITE).setBackground(CLR.TEAL)
     .setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  });
  return row + 1;
}

function tableHdrs12_(sh, row, labels) {
  rh_(sh, row, 26);
  mw_(sh, row, 2, 3).setValue(labels[0])
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.TEAL)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  for (let i = 2; i < labels.length; i++) {
    const c = sh.getRange(row, i + 2);
    if (labels[i]) {
      c.setValue(labels[i])
       .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
       .setFontColor(CLR.WHITE).setBackground(CLR.TEAL)
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
    } else {
      c.setBackground(CLR.TEAL);
    }
  }
  return row + 1;
}

function borders_(sh, row, sc, ec) {
  sh.getRange(row, sc, 1, ec - sc + 1)
    .setBorder(true, true, true, true, true, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

function grandTotal_(sh, row, sc, ec, label, totF, pasF, faiF, incF) {
  rh_(sh, row, 26);
  mw_(sh, row, sc, sc + 1)
    .setValue(label)
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  [
    [sc+2, totF,                               CLR.NAVY,    CLR.GOLD,  null   ],
    [sc+3, pasF,                               CLR.GREEN_DK,CLR.WHITE, null   ],
    [sc+4, faiF,                               '#9C0006',   CLR.WHITE, null   ],
    [sc+5, incF,                               CLR.PURPLE,  CLR.WHITE, null   ],
    [sc+6, `IFERROR((${pasF})/(${totF}),"—")`, CLR.GREEN_DK,CLR.WHITE, '0.0%' ],
    [sc+7, `IFERROR((${faiF})/(${totF}),"—")`, '#9C0006',   CLR.WHITE, '0.0%' ],
  ].forEach(([col, f, bg, color, fmt]) => {
    const c = sh.getRange(row, col);
    c.setFormula(`=${f}`)
     .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
     .setFontColor(color).setBackground(bg)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (fmt) c.setNumberFormat(fmt);
  });
  borders_(sh, row, sc, ec);
  return row + 1;
}

function grandTotal12_(sh, row, totF, pasF, faiF, incF) {
  rh_(sh, row, 26);
  mw_(sh, row, 2, 3)
    .setValue('Grand Total')
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  [
    [4, totF,                                  CLR.NAVY,    CLR.GOLD,  null   ],
    [5, pasF,                                  CLR.GREEN_DK,CLR.WHITE, null   ],
    [6, faiF,                                  '#9C0006',   CLR.WHITE, null   ],
    [7, incF,                                  CLR.PURPLE,  CLR.WHITE, null   ],
    [8, `IFERROR((${pasF})/(${totF}),"—")`,    CLR.GREEN_DK,CLR.WHITE, '0.0%' ],
    [9, `IFERROR((${faiF})/(${totF}),"—")`,    '#9C0006',   CLR.WHITE, '0.0%' ],
  ].forEach(([col, f, bg, color, fmt]) => {
    const c = sh.getRange(row, col);
    c.setFormula(`=${f}`)
     .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
     .setFontColor(color).setBackground(bg)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (fmt) c.setNumberFormat(fmt);
  });
  mw_(sh, row, 10, 12).setBackground(CLR.NAVY2);
  borders_(sh, row, 2, 9);
  return row + 1;
}

function grandTotalTemp_(sh, row, src, mc) {
  const t  = mfTotal_(src, mc, '');
  const p  = mfc_(src, mc, '', 'Pass');
  const f  = mfc_(src, mc, '', 'Fail');
  const ic = mfc_(src, mc, '', 'Inconclusive');
  rh_(sh, row, 26);
  mw_(sh, row, 2, 3)
    .setValue('Grand Total')
    .setFontFamily('Calibri').setFontSize(10).setFontWeight('bold')
    .setFontColor(CLR.WHITE).setBackground(CLR.NAVY2)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  [
    [4, t,                                CLR.NAVY,    CLR.GOLD,  null   ],
    [5, p,                                CLR.GREEN_DK,CLR.WHITE, null   ],
    [6, f,                                '#9C0006',   CLR.WHITE, null   ],
    [7, ic,                               CLR.PURPLE,  CLR.WHITE, null   ],
    [8, `IFERROR((${p})/(${t}),"—")`,     CLR.GREEN_DK,CLR.WHITE, '0.0%' ],
    [9, `IFERROR((${f})/(${t}),"—")`,     '#9C0006',   CLR.WHITE, '0.0%' ],
  ].forEach(([col, fm, bg, color, fmt]) => {
    const c = sh.getRange(row, col);
    c.setFormula(`=${fm}`)
     .setFontFamily('Calibri').setFontSize(11).setFontWeight('bold')
     .setFontColor(color).setBackground(bg)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (fmt) c.setNumberFormat(fmt);
  });
  cf_(sh, row, 10, `=${mfAvg_(src, mc, '')}`, CLR.TEAL, CLR.WHITE, true, '0.0');
  cf_(sh, row, 11, `=${mfMin_(src, mc, '')}`, CLR.TEAL, CLR.WHITE, false, '0.0');
  cf_(sh, row, 12, `=${mfMax_(src, mc, '')}`, CLR.TEAL, CLR.WHITE, false, '0.0');
  borders_(sh, row, 2, 12);
  return row + 1;
}

