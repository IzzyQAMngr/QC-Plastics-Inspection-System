/*************************************************************
 * PLASTICS QC INSPECTION SYSTEM V4.1 — FINAL
 * Container Supply Co. — Garden Grove, CA
 *
 * FORM LAYOUT:
 *   Row 1-2: Title / Subtitle
 *   Row 3:   Spacer
 *   Row 4:   Insp Date (B4), Insp Time (F4), Shift (J4),
 *            QC Record # (Z4), Save checkbox (AJ4)
 *   Row 5:   Inspector (B5), Foreman (F5), Clear checkbox (AJ5)
 *   Row 6-7: Info / spacer
 *   Row 8:   Column headers (A-AJ)
 *   Row 9+:  Data rows
 *
 * DATABASE (28 cols):
 *   A: QC Record #, B: Timestamp Saved, C: Inspection ID,
 *   D: Inspection Date, E: Inspection Time, F: Inspected By,
 *   G: Shift, H: Shift Foreman, I: Line #, J: Product Type,
 *   K: Mold, L: Color, M: LOT of Resin, N: Pallet Sequence,
 *   O: Sample Date, P: Sample Time, Q: Cavity ID,
 *   R: Test Type, S: Measure Index, T: Characteristic Name,
 *   U: Unit, V: LSL, W: USL, X: Actual Value,
 *   Y: Status, Z: Status Detail, AA: Visual Notes, AB: Source
 *
 * FEATURES:
 *   1.  LOAD checkbox — specs + cavities from Master Register
 *   2.  Blue spec rows (LSL/USL) + color specs (Type#, targets)
 *   3.  Soft lock (yellow) until Load checked
 *   4.  Red cell highlighting on out-of-spec
 *   5.  Wall thickness variance / eval auto-calc
 *   6.  Status column (AJ) — PASS or FAIL with detail
 *   7.  Date Code Verification — toast on Fail + included in email
 *   8.  Dynamic Mold dropdown filtered by Product Type
 *   9.  Dynamic Color dropdown filtered by Mold
 *   10. SAVE to tall-format database with Measure Index
 *   11. Email notification on failures (HTML + PDF)
 *   12. Column group shading (subtle tints)
 *   13. CLEAR form (restores shading, removes spec rows)
 *************************************************************/

// ================= SHEET NAMES =================
const FORM_SHEET     = "QC Form";
const DB_SHEET       = "QC Database";
const SETTINGS_SHEET = "Settings";

// ================= HEADER CELL REFERENCES =================
const SAVE_CELL  = "AJ4";
const CLEAR_CELL = "AJ5";

const HDR = {
  INSP_DATE:  "B4",
  INSP_TIME:  "F4",
  SHIFT:      "J4",
  INSPECTOR:  "B5",
  FOREMAN:    "F5",
  RECORD_ID:  "Z4",
};

const HDR_CLEAR_CELLS = ["B4","F4","J4","B5","F5","Z4"];

// ================= TABLE LAYOUT =================
const TABLE_START_ROW = 9;

const COL = {
  LOAD:1, LINE:2, PRODUCT:3, MOLD:4, COLOR:5, RESIN_LOT:6, PALLET:7,
  SAMPLE_DATE:8, SAMPLE_TIME:9, CAVITY:10, VISUAL:11, DATE_CODE:12,
  DELTA_L:13, DELTA_A:14, DELTA_B:15, DELTA_E:16,
  WEIGHT:17, BOT_THICK:18, SW1:19, SW2:20, SW3:21, SW4:22,
  SW_VAR:23, SW_EVAL:24,
  TOP_DIA:25, BOT_DIA:26, END_PIN:27, HEIGHT:28,
  CHIME_A_MIN:29, CHIME_A_MAX:30, CHIME_B_MIN:31, CHIME_B_MAX:32,
  NESTING:33, COVER_FIT:34, GAUGE_FIT:35, STATUS:36,
};
const TABLE_END_COL = 36;

const SPEC_CHECK_COLS = [17,18,19,20,21,22,25,26,27,28,29,30,31,32];
const WALL_COLS = [19,20,21,22];
const REPEATED_START = 2;
const REPEATED_END   = 9;
const MEASURE_START  = 11;

// ================= STYLING =================
const SPEC_ROW_BG   = "#dce6f1";
const SPEC_ROW_FONT = "#0c447c";
const OOS_BG        = "#f4cccc";
const OOS_FONT      = "#9c0006";
const SOFT_LOCK_BG  = "#fff2cc";
const PASS_BG       = "#d4edda";
const PASS_FONT     = "#0f6e56";

const GROUP_BG = {
  SAMPLE:      "#ffffff",
  VISUAL:      "#f9f9f9",
  COLOR:       "#f7f3ff",
  DIMENSIONAL: "#f3f8ff",
  FUNCTIONAL:  "#f3fff5",
  STATUS:      "#ffffff",
};

function getGroupBg_(col) {
  if (col <= 10) return GROUP_BG.SAMPLE;
  if (col <= 12) return GROUP_BG.VISUAL;
  if (col <= 16) return GROUP_BG.COLOR;
  if (col <= 32) return GROUP_BG.DIMENSIONAL;
  if (col <= 35) return GROUP_BG.FUNCTIONAL;
  return GROUP_BG.STATUS;
}

function applyGroupShadingRange_(sheet, row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    sheet.getRange(row, col).setBackground(getGroupBg_(col));
  }
}

// ================= MASTER REGISTER CONFIG =================
const PRODUCT_TO_TAB = { "pail":"Pail", "lid":"Cover", "cover":"Cover", "tub":"Tub", "handle":"Handle" };
const MR_DATA_START = 7;

const COL_CHAR_MAP = {
  17: { name:"Weight", match:["weight"] },
  18: { name:"Bottom Thickness", match:["bottom thickness"] },
  19: { name:"SW Thickness 1 - Top", match:["sw thickness 1","side wall thickness 1"] },
  20: { name:"SW Thickness 2 - Bottom", match:["sw thickness 2","side wall thickness 2"] },
  21: { name:"SW Thickness 3 - Side Right", match:["sw thickness 3","side wall thickness 3"] },
  22: { name:"SW Thickness 4 - Side Left", match:["sw thickness 4","side wall thickness 4"] },
  25: { name:"Top Diameter", match:["top diameter","top od"] },
  26: { name:"Bottom Diameter", match:["bottom diameter","bottom od"] },
  27: { name:"End Pin Diameter", match:["end pin"] },
  28: { name:"Height", match:["height"] },
  29: { name:"Chime A", match:["chime a"] },
  30: { name:"Chime A", match:["chime a"] },
  31: { name:"Chime B", match:["chime b"] },
  32: { name:"Chime B", match:["chime b"] },
};

const SHORT_NAMES = {
  17:"Wt", 18:"BotTh", 19:"SW1", 20:"SW2", 21:"SW3", 22:"SW4",
  25:"TopDia", 26:"BotDia", 27:"EndPin", 28:"Ht",
  29:"ChA.min", 30:"ChA.max", 31:"ChB.min", 32:"ChB.max",
};

const SETTINGS_EMAIL_CELL = "B3";
const SETTINGS_MR_ID_CELL = "B5";
const MOLD_LIST_COLS = { "pail":"D", "cover":"E", "tub":"F", "handle":"G" };


// ================= MENU =================
function onOpen() {
  ensureCheckboxes_();
  SpreadsheetApp.getUi().createMenu("QC Tools")
    .addItem("Set Master Register URL", "setMasterRegisterUrl")
    .addItem("Set Notification Emails", "setNotificationEmails")
    .addItem("Protect Form Header", "protectFormHeader_")
    .addSeparator()
    .addItem("Build Dashboard", "buildDashboard")
    .addItem("Build Shift Dashboard", "buildShiftDashboard")
    .addItem("Build Spec & Process Review", "buildProcessIntelligenceDashboard")
    .addItem("Set Spec Effective Date", "setSpecEffectiveDate")
    .addToUi();
}

function setMasterRegisterUrl() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Master Register Setup", "Paste the full URL or Spreadsheet ID:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  let input = resp.getResponseText().trim();
  const m = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) input = m[1];
  try {
    SpreadsheetApp.openById(input);
    PropertiesService.getScriptProperties().setProperty("MASTER_REG_ID", input);
    const st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
    if (st) st.getRange(SETTINGS_MR_ID_CELL).setValue(input);
    ui.alert("✅ Master Register linked!");
  } catch (e) { ui.alert("❌ Cannot open: " + e.message); }
}

function setNotificationEmails() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Notification Emails", "Enter comma-separated emails:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (st) st.getRange(SETTINGS_EMAIL_CELL).setValue(resp.getResponseText().trim());
  ui.alert("✅ Emails saved.");
}

function protectFormHeader_() {
  const form = SpreadsheetApp.getActive().getSheetByName(FORM_SHEET);
  if (!form) return;

  // ── Step 1: wipe ALL existing range protections on rows 1–8 ──────────────
  // (catches any sheet-UI protections, not just ones we created)
  for (const p of form.getProtections(SpreadsheetApp.ProtectionType.RANGE)) {
    try {
      const row = p.getRange().getRow();
      if (row >= 1 && row <= 8) p.remove();
    } catch (e) { /* skip if already removed */ }
  }

  // ── Step 2: rows 1–3 and 6–8 fully locked (title / col headers) ──────────
  form.getRange(1, 1, 3, TABLE_END_COL)
    .protect().setDescription("Header lock: title").setWarningOnly(false);
  form.getRange(6, 1, 3, TABLE_END_COL)
    .protect().setDescription("Header lock: col headers").setWarningOnly(false);

  // ── Step 3: row 4 — lock everything EXCEPT B4, F4, J4, AJ4 ───────────────
  //   Locked segments: A4 | C4:E4 | G4:I4 | K4:AI4
  //   Free (no protection at all): B4, F4, J4, AJ4
  form.getRange(4, 1, 1, 1)
    .protect().setDescription("Header lock: r4-A").setWarningOnly(false);
  form.getRange(4, 3, 1, 3)
    .protect().setDescription("Header lock: r4-C:E").setWarningOnly(false);
  form.getRange(4, 7, 1, 3)
    .protect().setDescription("Header lock: r4-G:I").setWarningOnly(false);
  form.getRange(4, 11, 1, 25)
    .protect().setDescription("Header lock: r4-K:AI").setWarningOnly(false);

  // ── Step 4: row 5 — lock everything EXCEPT B5, F5, AJ5 ───────────────────
  //   Locked segments: A5 | C5:E5 | G5:AI5
  //   Free (no protection at all): B5, F5, AJ5
  form.getRange(5, 1, 1, 1)
    .protect().setDescription("Header lock: r5-A").setWarningOnly(false);
  form.getRange(5, 3, 1, 3)
    .protect().setDescription("Header lock: r5-C:E").setWarningOnly(false);
  form.getRange(5, 7, 1, 29)
    .protect().setDescription("Header lock: r5-G:AI").setWarningOnly(false);

  SpreadsheetApp.getActive().toast(
    "B4, F4, J4, B5, F5 are now freely editable — all other header rows are locked.",
    "✅ Header Protected"
  );
}

// ================= ON EDIT TRIGGER =================
function qcOnEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== FORM_SHEET) return;

  const row = e.range.getRow(), col = e.range.getColumn();
  const a1 = e.range.getA1Notation(), val = e.range.getValue();

  // SAVE
  if (a1 === SAVE_CELL && val === true) {
    try { saveQcRecord_(); toast_("✅ QC record saved.", "SUCCESS"); }
    catch (err) { toast_("Save failed: " + err.message, "ERROR"); }
    finally { sheet.getRange(SAVE_CELL).setValue(false); }
    return;
  }

  // CLEAR
  if (a1 === CLEAR_CELL && val === true) {
    try { clearForm_(sheet); toast_("Form cleared.", "Cleared"); }
    catch (err) { toast_("Clear failed: " + err.message, "ERROR"); }
    finally { sheet.getRange(CLEAR_CELL).setValue(false); }
    return;
  }

  if (row < TABLE_START_ROW || isSpecRow_(sheet, row)) return;

  // LOAD
  if (col === COL.LOAD && val === true) {
    try { loadForRow_(sheet, row); toast_("Specs & cavities loaded.", "✅ Loaded"); }
    catch (err) {
      toast_("Load failed: " + err.message, "ERROR");
      // Only uncheck if the row is still a data row (not shifted to spec row)
      if (!isSpecRow_(sheet, row)) sheet.getRange(row, COL.LOAD).setValue(false);
    }
    return;
  }

  // DATE CODE VERIFICATION — toast on Fail
  if (col === COL.DATE_CODE) {
    if (String(val).toLowerCase() === "fail") {
      const mold = gv_(sheet, row, COL.MOLD);
      const line = gv_(sheet, row, COL.LINE);
      toast_("⚠️ Date Code incorrect — notify production to update date code for Line " + line + " / Mold " + mold, "DATE CODE ALERT");
    }
    updateRowStatus_(sheet, row);
    return;
  }

  // PRODUCT TYPE → update mold dropdown + soft lock
  if (col === COL.PRODUCT) {
    updateMoldDropdown_(sheet, row);
    if (!hasSpecRowsAbove_(sheet, row)) applySoftLock_(sheet, row);
    return;
  }

  // MOLD → color dropdown + soft lock
  if (col === COL.MOLD) {
    updateColorDropdown_(sheet, row);
    if (!hasSpecRowsAbove_(sheet, row)) applySoftLock_(sheet, row);
    return;
  }

  // WALL THICKNESS
  if (WALL_COLS.indexOf(col) >= 0) {
    updateWallThickness_(sheet, row);
    updateSpecHighlight_(sheet, row, col);
    return;
  }

  // SPEC-CHECKED COLUMNS
  if (SPEC_CHECK_COLS.indexOf(col) >= 0) {
    updateSpecHighlight_(sheet, row, col);
    updateRowStatus_(sheet, row);
    return;
  }

  // FUNCTIONAL COLUMNS
  if (col >= COL.NESTING && col <= COL.GAUGE_FIT) {
    updateRowStatus_(sheet, row);
    return;
  }
}


// ================= DYNAMIC MOLD DROPDOWN =================
function updateMoldDropdown_(sheet, row) {
  const product = gv_(sheet, row, COL.PRODUCT).toLowerCase().trim();
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName(SETTINGS_SHEET);
  if (!st) return;

  const col = MOLD_LIST_COLS[product];
  if (!col) {
    sheet.getRange(row, COL.MOLD).clearDataValidations();
    return;
  }

  const lastRow = st.getLastRow();
  const molds = st.getRange(col + "2:" + col + lastRow).getValues()
    .flat().map(v => String(v).trim()).filter(v => v !== "");

  if (molds.length === 0) {
    sheet.getRange(row, COL.MOLD).clearDataValidations();
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(molds, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, COL.MOLD).setDataValidation(rule);
  sheet.getRange(row, COL.MOLD).clearContent();
}


// ================= DYNAMIC COLOR DROPDOWN =================
function updateColorDropdown_(sheet, row) {
  const mold = gv_(sheet, row, COL.MOLD);
  if (!mold) {
    sheet.getRange(row, COL.COLOR).clearDataValidations();
    return;
  }

  try {
    const mr = SpreadsheetApp.openById(getMasterRegisterId_());
    const tab = mr.getSheetByName("Color Testing");
    if (!tab) return;
    const lastRow = tab.getLastRow();
    if (lastRow < 7) return;
    const data = tab.getRange(7, 1, lastRow - 6, 14).getValues();
    const colors = [];
    for (const r of data) {
      if (String(r[1] || "").trim() === mold &&
          String(r[11] || "").toLowerCase().trim() === "active") {
        const c = String(r[3] || "").trim();
        if (c && colors.indexOf(c) < 0) colors.push(c);
      }
    }
    if (colors.length === 0) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(colors, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, COL.COLOR).setDataValidation(rule);
  } catch (e) { /* silently fail — color dropdown is optional */ }
}


// ================= MASTER REGISTER ACCESS =================
function getMasterRegisterId_() {
  let id = PropertiesService.getScriptProperties().getProperty("MASTER_REG_ID");
  if (id) return id;
  const st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (st) {
    id = String(st.getRange(SETTINGS_MR_ID_CELL).getValue() || "").trim();
    if (id) {
      PropertiesService.getScriptProperties().setProperty("MASTER_REG_ID", id);
      return id;
    }
  }
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Master Register Setup", "Paste the URL or Spreadsheet ID:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) throw new Error("Master Register URL required.");
  let input = resp.getResponseText().trim();
  const m = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) input = m[1];
  try {
    SpreadsheetApp.openById(input);
    PropertiesService.getScriptProperties().setProperty("MASTER_REG_ID", input);
    if (st) st.getRange(SETTINGS_MR_ID_CELL).setValue(input);
    return input;
  } catch (e) { throw new Error("Cannot open Master Register: " + e.message); }
}

function getSpecsFromMaster_(productType, mold) {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tabName = PRODUCT_TO_TAB[String(productType).toLowerCase().trim()];
  if (!tabName) throw new Error("Unknown product type: " + productType);
  const tab = mr.getSheetByName(tabName);
  if (!tab) throw new Error("Tab '" + tabName + "' not found in Master Register.");
  const lastRow = tab.getLastRow();
  if (lastRow < MR_DATA_START) return [];
  const data = tab.getRange(MR_DATA_START, 1, lastRow - MR_DATA_START + 1, 8).getValues();
  const target = String(mold).trim();
  const results = [];
  for (const row of data) {
    if (String(row[0]).trim() !== target) continue;
    const lsl = row[5], usl = row[7], nom = row[6];
    results.push({
      characteristic: String(row[3] || "").trim(),
      unit: String(row[4] || "").trim(),
      measureIndex: row[2] !== "" && row[2] !== null ? row[2] : "",
      lsl: (lsl === "" || lsl === null) ? null : parseFloat(lsl),
      nominal: (nom === "" || nom === null) ? null : parseFloat(nom),
      usl: (usl === "" || usl === null) ? null : parseFloat(usl),
    });
  }
  return results;
}

function getColorSpec_(mold, color) {
  const mr = SpreadsheetApp.openById(getMasterRegisterId_());
  const tab = mr.getSheetByName("Color Testing");
  if (!tab) return null;
  const lastRow = tab.getLastRow();
  if (lastRow < 7) return null;
  const data = tab.getRange(7, 1, lastRow - 6, 14).getValues();
  const targetMold = String(mold).trim();
  const targetColor = String(color).toLowerCase().trim();
  for (const row of data) {
    if (String(row[1] || "").trim() !== targetMold) continue;
    if (String(row[3] || "").toLowerCase().trim() !== targetColor) continue;
    if (String(row[11] || "").toLowerCase().trim() !== "active") continue;
    return {
      typeNumber: String(row[0] || "").trim(),
      color: String(row[3] || "").trim(),
      L: row[4], a: row[5], b: row[6],
      deltaEMax: String(row[7] || "").trim(),
    };
  }
  return null;
}

function getCavityIds_(mold) {
  // Reads from Master Register → All Molds tab
  // Cols: A=Category, B=Product Description, C=Mold #s, D=Cavitation, E=Cavity IDs
  try {
    const mr = SpreadsheetApp.openById(getMasterRegisterId_());
    const tab = mr.getSheetByName("All Molds");
    if (!tab) return [];
    const lastRow = tab.getLastRow();
    if (lastRow < 2) return [];
    const data = tab.getRange(2, 3, lastRow - 1, 3).getDisplayValues(); // C:E
    const targetMold = String(mold).trim();
    const ids = [];
    for (const row of data) {
      if (String(row[0]).trim() === targetMold) {
        const cav = String(row[2]).trim();
        if (cav) ids.push(cav);
      }
    }
    return ids;
  } catch (e) { return []; }
}

function findSpecForCol_(specs, colNum) {
  const mapping = COL_CHAR_MAP[colNum];
  if (!mapping) return null;
  for (const spec of specs) {
    const name = String(spec.characteristic).toLowerCase()
      .replace(/\(.*?\)/g,"").replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
    for (const term of mapping.match) {
      if (name === term || name.indexOf(term) === 0 || term.indexOf(name) === 0) return spec;
    }
  }
  return null;
}


// ================= LOAD =================
function loadForRow_(sheet, rowNum) {
  const mold = gv_(sheet, rowNum, COL.MOLD);
  const product = gv_(sheet, rowNum, COL.PRODUCT);
  const color = gv_(sheet, rowNum, COL.COLOR);
  if (!mold) throw new Error("Enter a Mold before clicking Load.");
  if (!product) throw new Error("Enter a Product Type before clicking Load.");
  if (hasSpecRowsAbove_(sheet, rowNum)) throw new Error("Specs already loaded. Clear form to reload.");

  const specs = getSpecsFromMaster_(product, mold);
  if (specs.length === 0) throw new Error("No specs found for " + product + " / " + mold);

  const colorSpec = color ? getColorSpec_(mold, color) : null;
  const cavityIds = getCavityIds_(mold);
  const numCav = Math.max(cavityIds.length, 1);
  const extraRows = numCav - 1;
  const repeatedVals = sheet.getRange(rowNum, REPEATED_START, 1, REPEATED_END - REPEATED_START + 1).getValues()[0];

  sheet.insertRowsBefore(rowNum, 2);
  const lslRow = rowNum, uslRow = rowNum + 1, dataStart = rowNum + 2;

  // Clear inherited checkbox validation on spec rows so we can write "LSL"/"USL"
  sheet.getRange(lslRow, 1, 1, TABLE_END_COL).clearDataValidations();
  sheet.getRange(uslRow, 1, 1, TABLE_END_COL).clearDataValidations();

  // LSL row
  sheet.getRange(lslRow, COL.LOAD).setValue("LSL");
  sheet.getRange(lslRow, COL.MOLD).setValue(mold + " (" + product + ") — LSL");
  fillSpecRow_(sheet, lslRow, specs, "lsl");
  if (colorSpec) {
    sheet.getRange(lslRow, COL.DELTA_L).setValue(colorSpec.typeNumber);
  }
  styleSpecRow_(sheet, lslRow);

  // USL row
  sheet.getRange(uslRow, COL.LOAD).setValue("USL");
  sheet.getRange(uslRow, COL.MOLD).setValue(mold + " (" + product + ") — USL");
  fillSpecRow_(sheet, uslRow, specs, "usl");
  if (colorSpec) {
    sheet.getRange(uslRow, COL.DELTA_L).setValue(colorSpec.color);
    sheet.getRange(uslRow, COL.DELTA_E).setValue(colorSpec.deltaEMax);
  }
  styleSpecRow_(sheet, uslRow);

  // Lock spec rows so techs can't overwrite
  protectSpecRow_(sheet, lslRow, mold + " LSL");
  protectSpecRow_(sheet, uslRow, mold + " USL");

  if (extraRows > 0) {
    sheet.insertRowsAfter(dataStart, extraRows);
    sheet.getRange(dataStart, 1, 1, TABLE_END_COL)
      .copyTo(sheet.getRange(dataStart + 1, 1, extraRows, TABLE_END_COL));
  }

  for (let i = 0; i < numCav; i++) {
    const r = dataStart + i;
    sheet.getRange(r, REPEATED_START, 1, repeatedVals.length).setValues([repeatedVals]);
    if (cavityIds.length > 0) sheet.getRange(r, COL.CAVITY).setValue(cavityIds[i]);
    sheet.getRange(r, COL.LOAD).setValue(false);
    removeSoftLock_(sheet, r);
  }
}

function fillSpecRow_(sheet, row, specs, type) {
  for (const colStr in COL_CHAR_MAP) {
    const colNum = parseInt(colStr);
    const spec = findSpecForCol_(specs, colNum);
    if (!spec) continue;
    const val = type === "lsl" ? spec.lsl : spec.usl;
    if (val !== null && !isNaN(val)) sheet.getRange(row, colNum).setValue(val);
  }
}

function styleSpecRow_(sheet, row) {
  sheet.getRange(row, 1, 1, TABLE_END_COL)
    .setBackground(SPEC_ROW_BG).setFontColor(SPEC_ROW_FONT)
    .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
}

function protectSpecRow_(sheet, row, description) {
  const protection = sheet.getRange(row, 1, 1, TABLE_END_COL).protect();
  protection.setDescription("Spec row: " + description);
  protection.setWarningOnly(true);
}

function removeAllSpecProtections_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const p of protections) {
    if (p.getDescription().indexOf("Spec row:") === 0) {
      p.remove();
    }
  }
}


// ================= SPEC HELPERS =================
function isSpecRow_(sheet, row) {
  const v = String(sheet.getRange(row, COL.LOAD).getValue()).trim();
  return v === "LSL" || v === "USL";
}

function hasSpecRowsAbove_(sheet, row) {
  const myMold = gv_(sheet, row, COL.MOLD);
  if (!myMold) return false;
  for (let r = row - 1; r >= TABLE_START_ROW; r--) {
    if (isSpecRow_(sheet, r)) {
      // Check if this spec row belongs to our mold
      const specLabel = gv_(sheet, r, COL.MOLD);
      if (specLabel.indexOf(myMold) === 0) return true;
      return false; // Spec row for a different mold — not ours
    }
    // If we hit a data row with a different mold, stop looking
    const rowMold = gv_(sheet, r, COL.MOLD);
    if (rowMold && rowMold !== myMold) return false;
  }
  return false;
}

function findSpecRows_(sheet, row) {
  const myMold = gv_(sheet, row, COL.MOLD);
  let lslRow = null, uslRow = null;
  for (let r = row - 1; r >= TABLE_START_ROW; r--) {
    const marker = String(sheet.getRange(r, COL.LOAD).getValue()).trim();
    if (marker === "USL") { uslRow = r; continue; }
    if (marker === "LSL") { lslRow = r; break; }
    // If we hit a data row with a different mold, stop
    if (!isSpecRow_(sheet, r)) {
      const rowMold = gv_(sheet, r, COL.MOLD);
      if (rowMold && rowMold !== myMold) break;
    }
  }
  return { lslRow, uslRow };
}


// ================= SOFT LOCK =================
function applySoftLock_(sheet, row) {
  sheet.getRange(row, MEASURE_START, 1, TABLE_END_COL - MEASURE_START + 1).setBackground(SOFT_LOCK_BG);
}

function removeSoftLock_(sheet, row) {
  applyGroupShadingRange_(sheet, row, MEASURE_START, TABLE_END_COL);
}


// ================= SPEC HIGHLIGHTING =================
function updateSpecHighlight_(sheet, row, col) {
  if (isSpecRow_(sheet, row) || SPEC_CHECK_COLS.indexOf(col) < 0) return;
  const cell = sheet.getRange(row, col);
  const raw = cell.getValue();
  if (String(cell.getDisplayValue()).trim() === "" || isNaN(parseFloat(raw))) {
    cell.setBackground(getGroupBg_(col)).setFontColor("#000000").setFontWeight("normal");
    return;
  }
  const value = parseFloat(raw);
  const { lslRow, uslRow } = findSpecRows_(sheet, row);
  if (!lslRow || !uslRow) return;
  const lsl = parseFloat(sheet.getRange(lslRow, col).getValue());
  const usl = parseFloat(sheet.getRange(uslRow, col).getValue());
  const oos = (!isNaN(lsl) && value < lsl) || (!isNaN(usl) && value > usl);
  if (oos) cell.setBackground(OOS_BG).setFontColor(OOS_FONT).setFontWeight("bold");
  else cell.setBackground(getGroupBg_(col)).setFontColor("#000000").setFontWeight("normal");
}


// ================= WALL THICKNESS =================
function updateWallThickness_(sheet, row) {
  const vals = WALL_COLS.map(c => sheet.getRange(row, c).getValue());
  const trimmed = vals.map(v => String(v).trim());
  if (trimmed.every(v => v === "") || trimmed.some(v => v === "")) {
    sheet.getRange(row, COL.SW_VAR).clearContent();
    sheet.getRange(row, COL.SW_EVAL).clearContent();
    return;
  }
  const nums = vals.map(v => parseFloat(v));
  if (nums.some(v => isNaN(v))) {
    sheet.getRange(row, COL.SW_VAR).clearContent();
    sheet.getRange(row, COL.SW_EVAL).clearContent();
    return;
  }
  const variance = Math.max(...nums) - Math.min(...nums);
  sheet.getRange(row, COL.SW_VAR).setValue(Math.round(variance * 10000) / 10000);
  sheet.getRange(row, COL.SW_EVAL).setValue(variance < 0.005 ? "Pass" : "Fail");
  updateRowStatus_(sheet, row);
}


// ================= STATUS COLUMN =================
function updateRowStatus_(sheet, row) {
  if (isSpecRow_(sheet, row)) return;
  const { lslRow, uslRow } = findSpecRows_(sheet, row);
  const failures = [];

  const dateCode = gv_(sheet, row, COL.DATE_CODE).toLowerCase();
  if (dateCode === "fail") failures.push("Date Code Fail");

  for (const c of SPEC_CHECK_COLS) {
    const raw = sheet.getRange(row, c).getValue();
    if (String(sheet.getRange(row, c).getDisplayValue()).trim() === "" || isNaN(parseFloat(raw))) continue;
    const value = parseFloat(raw);
    if (lslRow && uslRow) {
      const lsl = parseFloat(sheet.getRange(lslRow, c).getValue());
      const usl = parseFloat(sheet.getRange(uslRow, c).getValue());
      const name = SHORT_NAMES[c] || ("Col" + c);
      if (!isNaN(lsl) && value < lsl) failures.push(name + " " + value + " < LSL " + lsl);
      if (!isNaN(usl) && value > usl) failures.push(name + " " + value + " > USL " + usl);
    }
  }

  const swEval = gv_(sheet, row, COL.SW_EVAL).toLowerCase();
  if (swEval === "fail") failures.push("SW Var. " + sheet.getRange(row, COL.SW_VAR).getValue() + " ≥ 0.005");

  for (const [c, name] of [[COL.NESTING,"Nesting"],[COL.COVER_FIT,"Cover Fit"],[COL.GAUGE_FIT,"Gauge Fit"]]) {
    const v = gv_(sheet, row, c).toLowerCase();
    if (v === "fail") failures.push(name + " Fail");
  }

  const statusCell = sheet.getRange(row, COL.STATUS);
  if (failures.length > 0) {
    statusCell.setValue("FAIL: " + failures.join("; "));
    statusCell.setFontColor(OOS_FONT).setFontWeight("bold").setBackground(OOS_BG);
  } else {
    const hasData = SPEC_CHECK_COLS.some(c => gv_(sheet, row, c) !== "") ||
      gv_(sheet, row, COL.DATE_CODE) !== "" ||
      gv_(sheet, row, COL.NESTING) !== "" ||
      gv_(sheet, row, COL.COVER_FIT) !== "" ||
      gv_(sheet, row, COL.GAUGE_FIT) !== "";
    if (hasData) {
      statusCell.setValue("PASS");
      statusCell.setFontColor(PASS_FONT).setFontWeight("bold").setBackground(PASS_BG);
    } else {
      statusCell.clearContent().setBackground(getGroupBg_(COL.STATUS))
        .setFontColor("#000000").setFontWeight("normal");
    }
  }
}


// ================= SAVE =================
function saveQcRecord_() {
  const ss = SpreadsheetApp.getActive();
  const form = ss.getSheetByName(FORM_SHEET);
  const db = ss.getSheetByName(DB_SHEET);
  if (!form || !db) throw new Error("Form or Database sheet not found.");

  const inspDate  = form.getRange(HDR.INSP_DATE).getDisplayValue();
  const inspTime  = form.getRange(HDR.INSP_TIME).getDisplayValue();
  const shift     = form.getRange(HDR.SHIFT).getDisplayValue();
  const inspector = form.getRange(HDR.INSPECTOR).getDisplayValue();
  const foreman   = form.getRange(HDR.FOREMAN).getDisplayValue();

  let recordID = form.getRange(HDR.RECORD_ID).getDisplayValue().trim();
  if (!/^QC-\d{8}-\d{3}$/.test(recordID)) {
    recordID = makeRecordID_(db);
    form.getRange(HDR.RECORD_ID).setValue(recordID);
  }

  const timestamp = new Date();
  const lastRow = form.getLastRow();
  if (lastRow < TABLE_START_ROW) throw new Error("No data rows.");

  const specCache = {};
  function getSpecsForRow(product, mold) {
    const key = product + "|" + mold;
    if (!specCache[key]) {
      try { specCache[key] = getSpecsFromMaster_(product, mold); }
      catch (e) { specCache[key] = []; }
    }
    return specCache[key];
  }

  const dbRows = [];
  const failRows = [];
  let inspectionId = 0;

  for (let r = TABLE_START_ROW; r <= lastRow; r++) {
    if (isSpecRow_(form, r)) continue;
    const mold = gv_(form, r, COL.MOLD);
    if (!mold) continue;
    inspectionId++;

    const product  = gv_(form, r, COL.PRODUCT);
    const color    = gv_(form, r, COL.COLOR);
    const resinLot = gv_(form, r, COL.RESIN_LOT);
    const pallet   = gv_(form, r, COL.PALLET);
    const sDate    = gv_(form, r, COL.SAMPLE_DATE);
    const sTime    = gv_(form, r, COL.SAMPLE_TIME);
    const cavity   = gv_(form, r, COL.CAVITY);
    const notes    = gv_(form, r, COL.VISUAL);
    const line     = gv_(form, r, COL.LINE);
    const rowStatus= gv_(form, r, COL.STATUS);
    const { lslRow, uslRow } = findSpecRows_(form, r);
    const specs = getSpecsForRow(product, mold);

    function dbRow(testType, measIdx, charName, unit, lsl, usl, actual, status, detail) {
      return [
        recordID, timestamp, inspectionId, inspDate, inspTime, inspector,
        shift, foreman, line, product, mold, color, resinLot, pallet,
        sDate, sTime, cavity, testType, measIdx, charName, unit, lsl, usl,
        actual, status, detail, notes, "QC Form V4.1"
      ];
    }

    function specVal(sr, c) { return sr ? form.getRange(sr, c).getValue() || "" : ""; }

    function evalSpec(v, l, u) {
      if (v === "" || isNaN(parseFloat(v))) return { s:"", d:"" };
      const val = parseFloat(v);
      const lv = (l !== "" && !isNaN(parseFloat(l))) ? parseFloat(l) : null;
      const uv = (u !== "" && !isNaN(parseFloat(u))) ? parseFloat(u) : null;
      if (lv !== null && val < lv) return { s:"Fail", d:val + " < LSL " + lv };
      if (uv !== null && val > uv) return { s:"Fail", d:val + " > USL " + uv };
      return { s:"Pass", d:"" };
    }

    function getMeasureIndex(charName) {
      const spec = specs.find(s => s.characteristic === charName);
      return spec ? spec.measureIndex : "";
    }

    // VISUAL
    const dateCode = gv_(form, r, COL.DATE_CODE);
    if (dateCode) {
      const dcStatus = dateCode.toLowerCase() === "fail" ? "Fail" : (dateCode.toLowerCase() === "pass" ? "Pass" : "");
      dbRows.push(dbRow("Visual", "", "Date Code Verification", "", "", "", dateCode, dcStatus,
        dcStatus === "Fail" ? "Date code incorrect" : ""));
    }
    if (notes) {
      dbRows.push(dbRow("Visual", "", "Visual Conformance Check", "", "", "", notes, "", ""));
    }

    // COLOR
    for (const [c, name] of [[COL.DELTA_L,"ΔL"],[COL.DELTA_A,"ΔA"],[COL.DELTA_B,"ΔB"],[COL.DELTA_E,"ΔE*ab"]]) {
      const v = gv_(form, r, c);
      if (v) dbRows.push(dbRow("Color", "", name, "", "", "", v, "", ""));
    }

    // DIMENSIONAL
    const dimDefs = [
      [COL.WEIGHT,      "Weight",                      "g"],
      [COL.BOT_THICK,   "Bottom Thickness",            "in"],
      [COL.SW1,         "SW Thickness 1 - Top",        "in"],
      [COL.SW2,         "SW Thickness 2 - Bottom",     "in"],
      [COL.SW3,         "SW Thickness 3 - Side Right", "in"],
      [COL.SW4,         "SW Thickness 4 - Side Left",  "in"],
      [COL.SW_VAR,      "Wall Thickness Variance",     "in"],
      [COL.SW_EVAL,     "Wall Thickness Evaluation",   ""],
      [COL.TOP_DIA,     "Top Diameter",                "in"],
      [COL.BOT_DIA,     "Bottom Diameter",             "in"],
      [COL.END_PIN,     "End Pin Diameter",            "in"],
      [COL.HEIGHT,      "Height",                      "in"],
      [COL.CHIME_A_MIN, "Chime A Min",                "in"],
      [COL.CHIME_A_MAX, "Chime A Max",                "in"],
      [COL.CHIME_B_MIN, "Chime B Min",                "in"],
      [COL.CHIME_B_MAX, "Chime B Max",                "in"],
    ];
    for (const [c, name, unit] of dimDefs) {
      const v = gv_(form, r, c);
      if (!v) continue;
      const l = specVal(lslRow, c), u = specVal(uslRow, c);
      const ev = evalSpec(v, l, u);
      const mi = getMeasureIndex(name);
      dbRows.push(dbRow("Dimensional", mi, name, unit, l, u, v, ev.s, ev.d));
    }

    // FUNCTIONAL
    for (const [c, name] of [[COL.NESTING,"Nesting"],[COL.COVER_FIT,"Cover Fit"],[COL.GAUGE_FIT,"Gauge Fit"]]) {
      const v = gv_(form, r, c);
      if (!v) continue;
      const s = v.toLowerCase();
      const st = (s==="pass"||s==="good"||s==="x") ? "Pass" : ((s==="fail"||s==="loose"||s==="tight") ? "Fail" : "");
      dbRows.push(dbRow("Functional", "", name, "", "", "", v, st, st === "Fail" ? name + " Fail" : ""));
    }

    if (rowStatus && rowStatus.toUpperCase().indexOf("FAIL") === 0) {
      failRows.push({ mold: mold, cavity: cavity, status: rowStatus, product: product, line: line });
    }
  }

  if (dbRows.length === 0) throw new Error("No data to save.");

  const startRow = Math.max(getLastDataRow_(db) + 1, 2);
  db.getRange(startRow, 1, dbRows.length, dbRows[0].length).setValues(dbRows);

  if (failRows.length > 0) {
    try { sendFailNotification_(form, recordID, inspDate, inspector, shift, failRows); }
    catch (emailErr) { toast_("Data saved but email failed: " + emailErr.message, "⚠️ Warning"); }
  }

  clearForm_(form);
}


// ================= EMAIL NOTIFICATION =================
function sendFailNotification_(form, recordID, inspDate, inspector, shift, failRows) {
  const st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!st) return;
  const emailList = String(st.getRange(SETTINGS_EMAIL_CELL).getValue() || "").trim();
  if (!emailList) return;
  const emails = emailList.split(",").map(e => e.trim()).filter(e => e.indexOf("@") > 0);
  if (emails.length === 0) return;

  let html = '<div style="font-family:Calibri,sans-serif;max-width:700px;">';
  html += '<div style="background:#1A3A3A;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">';
  html += '<h2 style="margin:0;font-size:18px;">⚠️ QC FAIL NOTIFICATION</h2>';
  html += '<p style="margin:4px 0 0;font-size:13px;opacity:0.8;">Plastics QC Inspection System</p></div>';
  html += '<div style="background:#f8f9fa;padding:16px 20px;border:1px solid #dee2e6;">';
  html += '<table style="font-size:13px;margin-bottom:12px;">';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">QC Record #:</td><td><b>' + recordID + '</b></td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Date:</td><td>' + inspDate + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Inspector:</td><td>' + inspector + '</td></tr>';
  html += '<tr><td style="color:#666;padding:2px 12px 2px 0;">Shift:</td><td>' + shift + '</td></tr>';
  html += '</table>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
  html += '<tr style="background:#d9534f;color:white;">';
  html += '<th style="padding:8px;text-align:left;">Line</th>';
  html += '<th style="padding:8px;text-align:left;">Mold</th>';
  html += '<th style="padding:8px;text-align:left;">Cavity</th>';
  html += '<th style="padding:8px;text-align:left;">Failure Details</th></tr>';

  for (let i = 0; i < failRows.length; i++) {
    const f = failRows[i];
    const bg = i % 2 === 0 ? "#fff" : "#f8f8f8";
    const detail = f.status.replace(/^FAIL:\s*/i, "");
    html += '<tr style="background:' + bg + ';">';
    html += '<td style="padding:6px 8px;border-bottom:1px solid #eee;">' + (f.line || "") + '</td>';
    html += '<td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.mold + '</td>';
    html += '<td style="padding:6px 8px;border-bottom:1px solid #eee;">' + f.cavity + '</td>';
    html += '<td style="padding:6px 8px;border-bottom:1px solid #eee;color:#d9534f;font-weight:bold;">' + detail + '</td></tr>';
  }
  html += '</table></div>';
  html += '<div style="background:#eee;padding:10px 20px;border-radius:0 0 8px 8px;font-size:11px;color:#888;">';
  html += 'Sent automatically by QC Inspection System. PDF of the form is attached.</div></div>';

  const ss = SpreadsheetApp.getActive();
  const ssId = ss.getId();
  const sheetId = form.getSheetId();
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&gid=" + sheetId +
    "&size=letter&landscape=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=true";
  const token = ScriptApp.getOAuthToken();
  const pdfBlob = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } })
    .getBlob().setName(recordID + "_QC_Form.pdf");

  const moldList = [...new Set(failRows.map(f => f.mold))].join(", ");
  const subject = "⚠️ QC PLASTICS INSP. FAIL — " + recordID + " | " + moldList;

  for (const email of emails) {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: html, attachments: [pdfBlob] });
  }
}


// ================= CLEAR =================
function clearForm_(form) {
  HDR_CLEAR_CELLS.forEach(ref => form.getRange(ref).clearContent());

  const lastRow = form.getLastRow();
  if (lastRow < TABLE_START_ROW) return;

  // Remove spec row protections before deleting
  removeAllSpecProtections_(form);

  // Delete spec rows bottom-up
  for (let r = lastRow; r >= TABLE_START_ROW; r--) {
    if (isSpecRow_(form, r)) form.deleteRow(r);
  }

  const newLast = form.getLastRow();
  if (newLast < TABLE_START_ROW) return;

  const numRows = newLast - TABLE_START_ROW + 1;
  const range = form.getRange(TABLE_START_ROW, 1, numRows, TABLE_END_COL);

  // Clear values but preserve formulas (batch)
  const values = range.getValues();
  const formulas = range.getFormulas();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < TABLE_END_COL; c++) {
      if (!formulas[r][c]) values[r][c] = "";
    }
  }
  range.setValues(values);

  // Build background color array (batch instead of cell-by-cell)
  const bgRow = [];
  for (let c = 1; c <= TABLE_END_COL; c++) bgRow.push(getGroupBg_(c));
  const bgArray = [];
  for (let r = 0; r < numRows; r++) bgArray.push([...bgRow]);
  range.setBackgrounds(bgArray);

  // Reset font formatting (batch)
  range.setFontColor("#000000").setFontWeight("normal");

  // Restore checkboxes on Load column
  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const loadRange = form.getRange(TABLE_START_ROW, COL.LOAD, numRows, 1);
  loadRange.setDataValidation(cbRule);
  loadRange.setValue(false);

  // Restore Product Type dropdown
  const ptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pail","Cover","Tub","Handle"], true)
    .setAllowInvalid(false).build();
  form.getRange(TABLE_START_ROW, COL.PRODUCT, numRows, 1).setDataValidation(ptRule);

  // Restore Date Code / Nesting / Cover Fit / Gauge Fit dropdowns
  const pfRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pass","Fail","N/A"], true)
    .setAllowInvalid(false).build();
  form.getRange(TABLE_START_ROW, COL.DATE_CODE, numRows, 1).setDataValidation(pfRule);
  form.getRange(TABLE_START_ROW, COL.NESTING, numRows, 1).setDataValidation(pfRule);
  form.getRange(TABLE_START_ROW, COL.COVER_FIT, numRows, 1).setDataValidation(pfRule);
  form.getRange(TABLE_START_ROW, COL.GAUGE_FIT, numRows, 1).setDataValidation(pfRule);

  // Clear Mold and Color validations (dynamic — set fresh on Product/Mold selection)
  form.getRange(TABLE_START_ROW, COL.MOLD, numRows, 1).clearDataValidations();
  form.getRange(TABLE_START_ROW, COL.COLOR, numRows, 1).clearDataValidations();

  // Trim excess rows beyond 16 blank data rows
  const target = TABLE_START_ROW + 15;
  const cur = form.getLastRow();
  if (cur > target) form.deleteRows(target + 1, cur - target);

  // Restore Save/Clear checkboxes if accidentally deleted
  ensureCheckboxes_();
}


// ================= RECORD ID =================
function makeRecordID_(db) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const lastRow = getLastDataRow_(db);
  if (lastRow < 2) return "QC-" + today + "-001";
  const ids = db.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let max = 0;
  ids.forEach(id => {
    const m = String(id || "").match(/^QC-(\d{8})-(\d{3})$/);
    if (m && m[1] === today) { const s = Number(m[2]); if (s > max) max = s; }
  });
  return "QC-" + today + "-" + String(max + 1).padStart(3, "0");
}


// ================= UTILITIES =================
function gv_(sheet, row, col) {
  return String(sheet.getRange(row, col).getDisplayValue() || "").trim();
}

function toast_(msg, title) {
  SpreadsheetApp.getActive().toast(msg, title, 5);
}

// Finds the REAL last row with data in column A (prevents gap issue)
function getLastDataRow_(sheet) {
  const col = sheet.getRange("A1:A").getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (col[i][0] !== "" && col[i][0] !== null) return i + 1;
  }
  return 1;
}

// Ensures Save and Clear checkboxes exist — auto-heals if deleted
function ensureCheckboxes_() {
  const form = SpreadsheetApp.getActive().getSheetByName(FORM_SHEET);
  if (!form) return;
  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  const saveCell = form.getRange(SAVE_CELL);
  if (saveCell.getDataValidation() === null || String(saveCell.getValue()) === "") {
    saveCell.setDataValidation(cbRule).setValue(false);
  }

  const clearCell = form.getRange(CLEAR_CELL);
  if (clearCell.getDataValidation() === null || String(clearCell.getValue()) === "") {
    clearCell.setDataValidation(cbRule).setValue(false);
  }
}


/*************************************************************
 * OVERALL DASHBOARD — PLASTICS QC INSPECTION DASHBOARD
 * 
 * Order:
 *   1. KPIs
 *   2. Functional Test Summary
 *   3. Pass/Fail Rate by Mold
 *   4. Dimensional Statistics by Product Type
 *   5. Mold Drill-Down
 *   6. Top Failure Combinations
 *   7. Color Testing (by color, grouped L/A/B/dE)
 *   8. Performance by Inspector
 *   9. Pass/Fail by Shift
 *************************************************************/

function buildDashboard() {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  if (!db) { SpreadsheetApp.getUi().alert("QC Database not found."); return; }

  const existing = ss.getSheetByName("📊 Dashboard");
  if (existing) ss.deleteSheet(existing);

  const dash = ss.insertSheet("📊 Dashboard");
  dash.setTabColor("#E67E22");

  const dbLast = db.getLastRow();
  if (dbLast < 2) { SpreadsheetApp.getUi().alert("No data in QC Database."); return; }

  // Gather unique values
  const allData = db.getRange(2, 1, dbLast - 1, 28).getValues();
  const moldsMap = {}, inspSet = new Set(), shiftsSet = new Set(), colorsSet = new Set();

  for (const row of allData) {
    const mold = String(row[10] || "").trim();
    const prod = String(row[9] || "").trim();
    const insp = String(row[5] || "").trim();
    const shift = String(row[6] || "").trim();
    const color = String(row[11] || "").trim();
    if (mold && prod) moldsMap[mold] = prod;
    if (insp) inspSet.add(insp);
    if (shift) shiftsSet.add(shift);
    if (color) colorsSet.add(color);
  }

  const productGroups = { "Pail": [], "Tub": [], "Cover": [], "Handle": [] };
  for (const [mold, prod] of Object.entries(moldsMap)) {
    if (productGroups[prod]) productGroups[prod].push(mold);
  }
  for (const key in productGroups) productGroups[key].sort();

  const allMolds = Object.keys(moldsMap).sort();
  const inspectors = [...inspSet].sort();
  const shifts = [...shiftsSet].sort((a, b) => Number(a) - Number(b));
  const colors = [...colorsSet].filter(c => c !== "").sort();

  const PAIL_CHARS = ["Weight","Bottom Thickness","SW Thickness 1 - Top","SW Thickness 2 - Bottom","SW Thickness 3 - Side Right","SW Thickness 4 - Side Left","Height","Chime A Min","Chime A Max","Chime B Min","Chime B Max"];
  const TUB_CHARS = ["Weight","Bottom Thickness","SW Thickness 1 - Top","SW Thickness 2 - Bottom","SW Thickness 3 - Side Right","SW Thickness 4 - Side Left","Height"];
  const COVER_CHARS = ["Weight","Wall Thickness","Bottom Thickness","Top Diameter","Bottom Diameter"];
  const HANDLE_CHARS = ["Weight","End Pin Diameter"];
  const ALL_DIM_CHARS = ["Weight","Bottom Thickness","SW Thickness 1 - Top","SW Thickness 2 - Bottom","SW Thickness 3 - Side Right","SW Thickness 4 - Side Left","Wall Thickness Variance","Height","Top Diameter","Bottom Diameter","End Pin Diameter","Chime A Min","Chime A Max","Chime B Min","Chime B Max"];

  const DARK = "#1A3A3A", MED = "#244747", COPPER = "#E67E22", WHITE = "#FFFFFF", LIGHT = "#F5F5F5";
  const DB = "'QC Database'";
  const maxCol = 20;

  function sectionHeader(r, text, endCol) {
    dash.getRange(r, 1, 1, endCol).merge().setValue(text)
      .setBackground(MED).setFontColor(WHITE).setFontSize(12).setFontWeight("bold")
      .setVerticalAlignment("middle");
    for (let c = 1; c <= endCol; c++) dash.getRange(r, c).setBackground(MED);
  }

  function tableHeader(r, headers, startCol) {
    for (let i = 0; i < headers.length; i++) {
      dash.getRange(r, startCol + i).setValue(headers[i])
        .setBackground(COPPER).setFontColor(WHITE).setFontWeight("bold").setFontSize(9)
        .setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  function dataCell(r, c, formula, fmt, bg) {
    const cell = dash.getRange(r, c);
    if (String(formula).charAt(0) === "=") cell.setFormula(formula);
    else cell.setValue(formula);
    cell.setFontSize(10).setHorizontalAlignment("center").setBackground(bg || WHITE)
      .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    if (fmt) cell.setNumberFormat(fmt);
    return cell;
  }

  function labelCell(r, c, text, bg) {
    dash.getRange(r, c).setValue(text).setFontSize(10).setBackground(bg || WHITE)
      .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  }

  // ==================== TITLE ====================
  let r = 1;
  dash.setRowHeight(1, 44);
  dash.getRange(1, 1, 1, maxCol).merge().setValue("📊  PLASTICS QC INSPECTION DASHBOARD")
    .setBackground(DARK).setFontColor(WHITE).setFontSize(18).setFontWeight("bold")
    .setVerticalAlignment("middle");

  dash.setRowHeight(2, 24);
  dash.getRange(2, 1, 1, maxCol).merge().setValue("Container Supply Co. — Garden Grove, CA")
    .setBackground(DARK).setFontColor(COPPER).setFontSize(13).setFontWeight("bold")
    .setVerticalAlignment("middle");

  // ==================== ROW 3: FILTERS ====================
  dash.setRowHeight(3, 32);
  for (let c = 1; c <= maxCol; c++) dash.getRange(3, c).setBackground(MED);
  dash.getRange(3, 1).setValue("START DATE").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(3, 2).setValue(new Date(2026, 0, 1)).setBackground("#0D2B2B").setFontColor(WHITE).setFontSize(10)
    .setNumberFormat("mm/dd/yyyy").setBorder(true,true,true,true,false,false,"#A8C4C4",SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(3, 4).setValue("END DATE").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(3, 5).setValue(new Date()).setBackground("#0D2B2B").setFontColor(WHITE).setFontSize(10)
    .setNumberFormat("mm/dd/yyyy").setBorder(true,true,true,true,false,false,"#A8C4C4",SpreadsheetApp.BorderStyle.SOLID);

  const SD = "$B$3", ED = "$E$3";

  // ==================== 1. KPIs ====================
  r = 5;
  sectionHeader(r, "KEY PERFORMANCE INDICATORS", maxCol);
  r++;
  dash.setRowHeight(r, 48); dash.setRowHeight(r+1, 18);
  for (let c = 1; c <= maxCol; c++) { dash.getRange(r, c).setBackground(DARK); dash.getRange(r+1, c).setBackground(DARK); }

  const kpis = [
    [1,2,`=COUNTA(UNIQUE(FILTER(${DB}!A2:A,${DB}!D2:D>=${SD},${DB}!D2:D<=${ED},${DB}!A2:A<>"")))`, "Total records", "0"],
    [4,5,`=IFERROR(COUNTIFS(${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})/COUNTIFS(${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),0)`, "Dim. pass rate", "0.0%"],
    [7,8,`=COUNTIFS(${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "Dim. failures", "0"],
    [10,11,`=COUNTIFS(${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "Dim. tests", "#,##0"],
    [13,14,`=COUNTA(UNIQUE(FILTER(${DB}!K2:K,${DB}!D2:D>=${SD},${DB}!D2:D<=${ED},${DB}!K2:K<>"")))`, "Active molds", "0"],
    [16,17,`=COUNTA(UNIQUE(FILTER(${DB}!F2:F,${DB}!D2:D>=${SD},${DB}!D2:D<=${ED},${DB}!F2:F<>"")))`, "Inspectors", "0"],
  ];
  for (const [c1,c2,formula,label,fmt] of kpis) {
    dash.getRange(r,c1,1,2).merge().setFormula(formula).setFontSize(22).setFontWeight("bold").setFontColor(WHITE)
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#244747").setNumberFormat(fmt)
      .setBorder(true,true,true,true,false,false,"#3A6363",SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r+1,c1,1,2).merge().setValue(label).setFontSize(9).setFontColor("#A8C4C4").setHorizontalAlignment("center").setBackground(DARK);
  }

  // ==================== 2. FUNCTIONAL TEST SUMMARY ====================
  r = 9;
  sectionHeader(r, "FUNCTIONAL TEST SUMMARY", maxCol);
  r++;
  tableHeader(r, ["Test","Total","Pass","Fail","Pass %"], 1);
  r++;
  const funcTests = ["Nesting","Cover Fit","Gauge Fit","Date Code Verification"];
  for (let i = 0; i < funcTests.length; i++) {
    const fr = r + i, test = funcTests[i], bg = i%2===1 ? LIGHT : WHITE;
    labelCell(fr, 1, test, bg);
    dataCell(fr, 2, `=COUNTIFS(${DB}!T:T,"${test}",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(fr, 3, `=COUNTIFS(${DB}!T:T,"${test}",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(fr, 4, `=COUNTIFS(${DB}!T:T,"${test}",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg).setFontColor("#9C0006");
    dataCell(fr, 5, `=IFERROR(C${fr}/B${fr},"")`, "0.0%", bg).setFontWeight("bold");
  }

  // ==================== 3. PASS/FAIL BY MOLD ====================
  r += funcTests.length + 2;
  sectionHeader(r, "PASS / FAIL RATE BY MOLD", maxCol);
  r++;
  tableHeader(r, ["Mold","Product","Tests","Pass","Fail","Pass %","Fail %"], 1);
  r++;
  for (let i = 0; i < allMolds.length; i++) {
    const mr = r + i, mold = allMolds[i], bg = i%2===1 ? LIGHT : WHITE;
    labelCell(mr, 1, mold, bg);
    dataCell(mr, 2, `=IFERROR(INDEX(${DB}!J:J,MATCH(A${mr},${DB}!K:K,0)),"")`, null, bg);
    dataCell(mr, 3, `=COUNTIFS(${DB}!K:K,A${mr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(mr, 4, `=COUNTIFS(${DB}!K:K,A${mr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(mr, 5, `=COUNTIFS(${DB}!K:K,A${mr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg).setFontColor("#9C0006");
    dataCell(mr, 6, `=IFERROR(D${mr}/C${mr},"")`, "0.0%", bg).setFontWeight("bold");
    dataCell(mr, 7, `=IFERROR(E${mr}/C${mr},"")`, "0.0%", bg).setFontWeight("bold").setFontColor("#9C0006");
  }

  // ==================== 4. DIMENSIONAL STATS BY PRODUCT TYPE ====================
  r += allMolds.length + 2;
  sectionHeader(r, "DIMENSIONAL STATISTICS BY PRODUCT TYPE", maxCol);
  r += 2;

  const MOLD_TINT_A = WHITE, MOLD_TINT_B = "#E8F0EF";
  const MOLD_HDR_A = "#244747", MOLD_HDR_B = "#2D5A5A";

  function buildProductTable(startRow, productName, molds, chars) {
    if (molds.length === 0) return startRow;
    let cr = startRow;
    dash.getRange(cr,1,1,maxCol).merge().setValue(productName + " — dimensional statistics")
      .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold");
    for (let c=1;c<=maxCol;c++) dash.getRange(cr,c).setBackground(COPPER);
    cr++;

    dash.getRange(cr,1).setValue("Characteristic").setBackground(MOLD_HDR_A).setFontColor(WHITE).setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    let col = 2;
    for (let mi=0; mi<molds.length; mi++) {
      dash.getRange(cr,col,1,3).merge().setValue(molds[mi]).setBackground(mi%2===0?MOLD_HDR_A:MOLD_HDR_B)
        .setFontColor(WHITE).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center")
        .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
      col+=3;
    }
    cr++;
    dash.getRange(cr,1).setValue("").setBackground(MED).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    col=2;
    for (let mi=0; mi<molds.length; mi++) {
      for (const lbl of ["Avg","Min","Max"]) {
        dash.getRange(cr,col).setValue(lbl).setBackground(mi%2===0?MOLD_HDR_A:MOLD_HDR_B).setFontColor("#A8C4C4")
          .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center")
          .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
        col++;
      }
    }
    cr++;
    for (let ci=0; ci<chars.length; ci++) {
      const charName = chars[ci], rowBg = ci%2===1?LIGHT:WHITE;
      const shortName = charName.replace("SW Thickness ","SW").replace(" - ","-").replace("Side Right","R").replace("Side Left","L")
        .replace("Wall Thickness Variance","SW Var").replace("Bottom Thickness","Bot Thick").replace("End Pin Diameter","End Pin")
        .replace("Top Diameter","Top Dia").replace("Bottom Diameter","Bot Dia")
        .replace("Chime A Min","ChA Min").replace("Chime A Max","ChA Max").replace("Chime B Min","ChB Min").replace("Chime B Max","ChB Max");
      labelCell(cr,1,shortName,rowBg); dash.getRange(cr,1).setFontWeight("bold");
      col=2;
      const fmt = charName==="Weight"?"0.0":"0.000";
      for (let mi=0; mi<molds.length; mi++) {
        const mold=molds[mi], moldBg = mi%2===0 ? rowBg : (ci%2===1?"#E0EAEA":MOLD_TINT_B);
        dataCell(cr,col,`=IFERROR(AVERAGEIFS(${DB}!X:X,${DB}!K:K,"${mold}",${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,moldBg);
        dataCell(cr,col+1,`=IFERROR(MINIFS(${DB}!X:X,${DB}!K:K,"${mold}",${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,moldBg);
        dataCell(cr,col+2,`=IFERROR(MAXIFS(${DB}!X:X,${DB}!K:K,"${mold}",${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,moldBg);
        col+=3;
      }
      cr++;
    }
    cr++;
    return cr;
  }

  r = buildProductTable(r, "PAILS", productGroups["Pail"], PAIL_CHARS);
  r = buildProductTable(r, "TUBS", productGroups["Tub"], TUB_CHARS);
  r = buildProductTable(r, "COVERS", productGroups["Cover"], COVER_CHARS);
  r = buildProductTable(r, "HANDLES", productGroups["Handle"], HANDLE_CHARS);

  // ==================== 5. MOLD DRILL-DOWN ====================
  r++;
  sectionHeader(r, "MOLD DRILL-DOWN — Select a mold to see detailed stats", maxCol);
  r++;
  for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(DARK);
  dash.getRange(r,1).setValue("SELECT MOLD →").setFontColor("#A8C4C4").setFontSize(11).setFontWeight("bold").setBackground(DARK);

  const moldDropCell = dash.getRange(r, 3);
  moldDropCell.setValue(allMolds[0] || "").setBackground(COPPER).setFontColor(WHITE).setFontSize(12).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true)
    .setBorder(true,true,true,true,false,false,WHITE,SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 3, 1, 2).merge(); // Wider for wrapped text
  const moldRule = SpreadsheetApp.newDataValidation().requireValueInList(allMolds, true).setAllowInvalid(false).build();
  dash.getRange(r, 3).setDataValidation(moldRule);
  dash.setRowHeight(r, 36);

  const MC = "$C$" + r;

  r += 2;
  for (let c=1;c<=maxCol;c++) { dash.getRange(r,c).setBackground(DARK); dash.getRange(r+1,c).setBackground(DARK); }
  dash.setRowHeight(r, 44);
  const dKpis = [
    [1,2,`=COUNTIFS(${DB}!K:K,${MC},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "Tests", "0"],
    [4,5,`=IFERROR(COUNTIFS(${DB}!K:K,${MC},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})/COUNTIFS(${DB}!K:K,${MC},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),0)`, "Pass rate", "0.0%"],
    [7,8,`=COUNTIFS(${DB}!K:K,${MC},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "Failures", "0"],
    [10,11,`=COUNTA(UNIQUE(FILTER(${DB}!Q2:Q,${DB}!K2:K=${MC},${DB}!Q2:Q<>"",${DB}!D2:D>=${SD},${DB}!D2:D<=${ED})))`, "Cavities", "0"],
  ];
  for (const [c1,c2,formula,label,fmt] of dKpis) {
    dash.getRange(r,c1,1,2).merge().setFormula(formula).setFontSize(20).setFontWeight("bold").setFontColor(WHITE)
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#244747").setNumberFormat(fmt)
      .setBorder(true,true,true,true,false,false,"#3A6363",SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r+1,c1,1,2).merge().setValue(label).setFontSize(9).setFontColor("#A8C4C4").setHorizontalAlignment("center").setBackground(DARK);
  }

  r += 3;
  // Dimensional stats
  dash.getRange(r,1,1,6).merge().setValue("Dimensional stats for selected mold")
    .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold");
  for (let c=1;c<=6;c++) dash.getRange(r,c).setBackground(COPPER);
  r++;
  tableHeader(r, ["Characteristic","Avg","Min","Max","LSL","USL"], 1);
  r++;
  const drillDimStart = r;

  for (let i=0; i<ALL_DIM_CHARS.length; i++) {
    const charName = ALL_DIM_CHARS[i], bg = i%2===1?LIGHT:WHITE;
    const fmt = charName==="Weight"?"0.0":"0.0000";
    const shortName = charName.replace("SW Thickness ","SW").replace(" - ","-").replace("Side Right","R").replace("Side Left","L")
      .replace("Wall Thickness Variance","SW Var").replace("Bottom Thickness","Bot Thick").replace("End Pin Diameter","End Pin")
      .replace("Top Diameter","Top Dia").replace("Bottom Diameter","Bot Dia")
      .replace("Chime A Min","ChA Min").replace("Chime A Max","ChA Max").replace("Chime B Min","ChB Min").replace("Chime B Max","ChB Max");
    labelCell(r,1,shortName,bg);
    dataCell(r,2,`=IFERROR(AVERAGEIFS(${DB}!X:X,${DB}!K:K,${MC},${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,bg);
    dataCell(r,3,`=IFERROR(MINIFS(${DB}!X:X,${DB}!K:K,${MC},${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,bg);
    dataCell(r,4,`=IFERROR(MAXIFS(${DB}!X:X,${DB}!K:K,${MC},${DB}!T:T,"${charName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`,fmt,bg);
    dataCell(r,5,`=IFERROR(INDEX(FILTER(${DB}!V:V,${DB}!K:K=${MC},${DB}!T:T="${charName}",${DB}!V:V<>""),1),"")`,fmt,bg).setFontColor("#0C447C");
    dataCell(r,6,`=IFERROR(INDEX(FILTER(${DB}!W:W,${DB}!K:K=${MC},${DB}!T:T="${charName}",${DB}!W:W<>""),1),"")`,fmt,bg).setFontColor("#0C447C");
    r++;
  }

  // Cavity breakdown (right side)
  const cavCol = 8;
  const cavHdrRow = drillDimStart - 2;
  dash.getRange(cavHdrRow,cavCol,1,5).merge().setValue("Cavity performance for selected mold")
    .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold");
  for (let c=cavCol;c<=cavCol+4;c++) dash.getRange(cavHdrRow,c).setBackground(COPPER);
  tableHeader(cavHdrRow+1, ["Cavity","Tests","Pass","Fail","Pass %"], cavCol);

  const maxCavRows = 20, cavDataStart = cavHdrRow + 2;
  dash.getRange(cavDataStart,cavCol).setFormula(
    `=IFERROR(SORT(UNIQUE(FILTER(${DB}!Q2:Q,${DB}!K2:K=${MC},${DB}!Q2:Q<>"",${DB}!D2:D>=${SD},${DB}!D2:D<=${ED}))),"")`
  ).setFontSize(10).setHorizontalAlignment("center")
    .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);

  for (let i=0; i<maxCavRows; i++) {
    const cr = cavDataStart + i, bg = i%2===1?LIGHT:WHITE;
    const cavRef = "$" + String.fromCharCode(64+cavCol) + "$" + cr;
    if (i > 0) dash.getRange(cr,cavCol).setFontSize(10).setHorizontalAlignment("center").setBackground(bg)
      .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    else dash.getRange(cr,cavCol).setBackground(bg);
    dataCell(cr,cavCol+1,`=IF(${cavRef}="","",COUNTIFS(${DB}!K:K,${MC},${DB}!Q:Q,${cavRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}))`, "#,##0", bg);
    dataCell(cr,cavCol+2,`=IF(${cavRef}="","",COUNTIFS(${DB}!K:K,${MC},${DB}!Q:Q,${cavRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}))`, "#,##0", bg);
    dataCell(cr,cavCol+3,`=IF(${cavRef}="","",COUNTIFS(${DB}!K:K,${MC},${DB}!Q:Q,${cavRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}))`, "#,##0", bg).setFontColor("#9C0006");
    const pc=String.fromCharCode(64+cavCol+2), tc=String.fromCharCode(64+cavCol+1);
    dataCell(cr,cavCol+4,`=IF(${cavRef}="","",IFERROR(${pc}${cr}/${tc}${cr},""))`, "0.0%", bg).setFontWeight("bold");
  }

  r = Math.max(r, cavDataStart + maxCavRows) + 2;

  // ==================== 6. TOP FAILURES ====================
  sectionHeader(r, "TOP FAILURE COMBINATIONS", maxCol);
  r++;
  tableHeader(r, ["Mold","Characteristic","Fail Count"], 1);
  r++;
  dash.getRange(r,1).setFormula(
    `=IFERROR(QUERY(${DB}!A:AB,"SELECT K, T, COUNT(T) WHERE R='Dimensional' AND Y='Fail' AND D >= date '"&TEXT(${SD},"yyyy-MM-dd")&"' AND D <= date '"&TEXT(${ED},"yyyy-MM-dd")&"' GROUP BY K, T ORDER BY COUNT(T) DESC LIMIT 15 LABEL K '', T '', COUNT(T) ''"),"")`
  ).setFontSize(10).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);

  r += 17;

  // ==================== 7. COLOR TESTING BY COLOR ====================
  const COLOR_GROUP_A = "#FFFFFF";
  const COLOR_GROUP_B = "#F0EBF8"; // Light purple tint for alternating groups

  function buildColorTable(startRow, colorName, molds) {
    let cr = startRow;
    dash.getRange(cr,1,1,maxCol).merge().setValue("Color: " + colorName)
      .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold");
    for (let c=1;c<=maxCol;c++) dash.getRange(cr,c).setBackground(COPPER);
    cr++;

    // Headers: Mold | ΔL (Avg Min Max) | ΔA (Avg Min Max) | ΔB (Avg Min Max) | ΔE* (Avg Min Max) | Readings
    dash.getRange(cr,1).setValue("Mold").setBackground(MOLD_HDR_A).setFontColor(WHITE).setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);

    const colorChars = ["ΔL","ΔA","ΔB","ΔE*ab"];
    let col = 2;
    for (let gi=0; gi<colorChars.length; gi++) {
      const gBg = gi%2===0 ? MOLD_HDR_A : MOLD_HDR_B;
      dash.getRange(cr,col,1,3).merge().setValue(colorChars[gi]).setBackground(gBg)
        .setFontColor(WHITE).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center")
        .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
      col += 3;
    }
    dash.getRange(cr,col).setValue("Readings").setBackground(MOLD_HDR_A).setFontColor(WHITE).setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    cr++;

    // Sub-headers
    dash.getRange(cr,1).setValue("").setBackground(MED).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    col = 2;
    for (let gi=0; gi<colorChars.length; gi++) {
      const gBg = gi%2===0 ? MOLD_HDR_A : MOLD_HDR_B;
      for (const lbl of ["Avg","Min","Max"]) {
        dash.getRange(cr,col).setValue(lbl).setBackground(gBg).setFontColor("#A8C4C4").setFontSize(9).setFontWeight("bold")
          .setHorizontalAlignment("center").setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
        col++;
      }
    }
    dash.getRange(cr,col).setValue("").setBackground(MED).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    cr++;

    // Data rows per mold
    for (let mi=0; mi<molds.length; mi++) {
      const mold = molds[mi], bg = mi%2===1 ? LIGHT : WHITE;
      labelCell(cr,1,mold,bg);
      col = 2;
      for (let gi=0; gi<colorChars.length; gi++) {
        const dbName = colorChars[gi];
        const gTint = gi%2===0 ? bg : (mi%2===1 ? "#E6E0EE" : COLOR_GROUP_B);
        dataCell(cr,col,`=IFERROR(AVERAGEIFS(${DB}!X:X,${DB}!K:K,$A${cr},${DB}!T:T,"${dbName}",${DB}!L:L,"${colorName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`, "0.00", gTint);
        dataCell(cr,col+1,`=IFERROR(MINIFS(${DB}!X:X,${DB}!K:K,$A${cr},${DB}!T:T,"${dbName}",${DB}!L:L,"${colorName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`, "0.00", gTint);
        dataCell(cr,col+2,`=IFERROR(MAXIFS(${DB}!X:X,${DB}!K:K,$A${cr},${DB}!T:T,"${dbName}",${DB}!L:L,"${colorName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED}),"")`, "0.00", gTint);
        col += 3;
      }
      dataCell(cr,col,`=COUNTIFS(${DB}!K:K,$A${cr},${DB}!T:T,"ΔE*ab",${DB}!L:L,"${colorName}",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
      cr++;
    }
    cr++;
    return cr;
  }

  sectionHeader(r, "COLOR TESTING BY COLOR", maxCol);
  r += 2;
  for (const color of colors) {
    r = buildColorTable(r, color, allMolds);
  }

  // ==================== 8. PERFORMANCE BY INSPECTOR ====================
  r++;
  sectionHeader(r, "PERFORMANCE BY INSPECTOR", maxCol);
  r++;
  tableHeader(r, ["Inspector","Tests","Pass","Fail","Pass %","Fail %","Visual"], 1);
  r++;
  for (let i=0; i<inspectors.length; i++) {
    const ir = r+i, bg = i%2===1?LIGHT:WHITE;
    labelCell(ir,1,inspectors[i],bg);
    dataCell(ir,2,`=COUNTIFS(${DB}!F:F,A${ir},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(ir,3,`=COUNTIFS(${DB}!F:F,A${ir},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(ir,4,`=COUNTIFS(${DB}!F:F,A${ir},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg).setFontColor("#9C0006");
    dataCell(ir,5,`=IFERROR(C${ir}/B${ir},"")`, "0.0%", bg).setFontWeight("bold");
    dataCell(ir,6,`=IFERROR(D${ir}/B${ir},"")`, "0.0%", bg).setFontWeight("bold").setFontColor("#9C0006");
    dash.getRange(ir,7).setFormula(`=SPARKLINE({C${ir},D${ir}},{"charttype","bar";"color1","#27AE60";"color2","#E74C3C"})`).setBackground(bg)
      .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  }

  // ==================== 9. PASS/FAIL BY SHIFT ====================
  r += inspectors.length + 2;
  sectionHeader(r, "PASS / FAIL BY SHIFT", maxCol);
  r++;
  tableHeader(r, ["Shift","Tests","Pass","Fail","Pass %","Fail %","Visual"], 1);
  r++;
  for (let i=0; i<shifts.length; i++) {
    const sr = r+i, bg = i%2===1?LIGHT:WHITE;
    labelCell(sr,1,shifts[i],bg); dash.getRange(sr,1).setHorizontalAlignment("center");
    dataCell(sr,2,`=COUNTIFS(${DB}!G:G,A${sr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(sr,3,`=COUNTIFS(${DB}!G:G,A${sr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg);
    dataCell(sr,4,`=COUNTIFS(${DB}!G:G,A${sr},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!D:D,">="&${SD},${DB}!D:D,"<="&${ED})`, "#,##0", bg).setFontColor("#9C0006");
    dataCell(sr,5,`=IFERROR(C${sr}/B${sr},"")`, "0.0%", bg).setFontWeight("bold");
    dataCell(sr,6,`=IFERROR(D${sr}/B${sr},"")`, "0.0%", bg).setFontWeight("bold").setFontColor("#9C0006");
    dash.getRange(sr,7).setFormula(`=SPARKLINE({C${sr},D${sr}},{"charttype","bar";"color1","#27AE60";"color2","#E74C3C"})`).setBackground(bg)
      .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  }

  // ==================== COLUMN WIDTHS ====================
  dash.setColumnWidth(1, 170);
  for (let c = 2; c <= 50; c++) dash.setColumnWidth(c, 85);

  dash.setFrozenRows(3);
  toast_("📊 Dashboard built!", "Done");
}



/*************************************************************
 * SHIFT PERFORMANCE DASHBOARD
 * Auto-detects the most recent QC record per shift.
 *************************************************************/

function buildShiftDashboard() {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  if (!db) { SpreadsheetApp.getUi().alert("QC Database not found."); return; }

  const existing = ss.getSheetByName("📋 Shift Performance");
  if (existing) ss.deleteSheet(existing);

  const dash = ss.insertSheet("📋 Shift Performance");
  dash.setTabColor("#E67E22");

  const DARK = "#1A3A3A", MED = "#244747", COPPER = "#E67E22", WHITE = "#FFFFFF", LIGHT = "#F5F5F5";
  const SHIFT_COLORS = { 1: "#27AE60", 2: "#2980B9", 3: "#8E44AD" };
  const DB = "'QC Database'";
  const maxCol = 14;
  const lines = [1,2,3,4,5,6,7,8,9,10];
  const refCol = maxCol + 1;

  function sectionHeader(r, text, endCol) {
    dash.getRange(r,1,1,endCol).merge().setValue(text)
      .setBackground(MED).setFontColor(WHITE).setFontSize(12).setFontWeight("bold").setVerticalAlignment("middle");
    for (let c=1;c<=endCol;c++) dash.getRange(r,c).setBackground(MED);
  }
  function tableHeader(r, headers, startCol) {
    for (let i=0;i<headers.length;i++) {
      dash.getRange(r,startCol+i).setValue(headers[i])
        .setBackground(COPPER).setFontColor(WHITE).setFontWeight("bold").setFontSize(9)
        .setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  function dataCell(r,c,formula,fmt,bg) {
    const cell = dash.getRange(r,c);
    if (String(formula).charAt(0)==="=") cell.setFormula(formula); else cell.setValue(formula);
    cell.setFontSize(10).setHorizontalAlignment("center").setBackground(bg||WHITE)
      .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    if (fmt) cell.setNumberFormat(fmt);
    return cell;
  }
  function labelCell(r,c,text,bg) {
    dash.getRange(r,c).setValue(text).setFontSize(10).setBackground(bg||WHITE)
      .setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  }

  // ==================== TITLE ====================
  dash.setRowHeight(1, 44);
  dash.getRange(1,1,1,maxCol).merge().setValue("📋  SHIFT PERFORMANCE DASHBOARD")
    .setBackground(DARK).setFontColor(WHITE).setFontSize(18).setFontWeight("bold").setVerticalAlignment("middle");

  dash.setRowHeight(2, 24);
  dash.getRange(2,1,1,maxCol).merge().setValue("Auto-populated with most recent QC check per shift — Container Supply Co.")
    .setBackground(DARK).setFontColor(COPPER).setFontSize(13).setFontWeight("bold").setVerticalAlignment("middle");

  dash.setRowHeight(3, 28);
  for (let c=1;c<=maxCol;c++) dash.getRange(3,c).setBackground(DARK);
  dash.getRange(3,1,1,maxCol).merge()
    .setValue("Each shift section shows the last saved QC record. Bars and trend lines update automatically.")
    .setFontColor("#A8C4C4").setFontSize(10).setFontStyle("italic").setBackground(DARK);

  // ==================== SHIFT SECTIONS ====================
  let r = 5;

  for (const shift of [1, 2, 3]) {
    const shiftColor = SHIFT_COLORS[shift];

    // Hidden helper: most recent QC Record # for this shift (all dates)
    dash.getRange(r, refCol).setFormula(
      `=IFERROR(INDEX(SORT(FILTER(${DB}!A2:A,${DB}!G2:G*1=${shift},${DB}!A2:A<>""),1,FALSE),1),"")`
    ).setFontSize(1).setFontColor(WHITE);
    const LAST_REC = "$" + String.fromCharCode(64 + refCol) + "$" + r;

    // Shift title
    dash.getRange(r,1,1,maxCol).merge().setBackground(shiftColor).setFontColor(WHITE).setFontSize(14).setFontWeight("bold").setVerticalAlignment("middle");
    dash.getRange(r,1).setValue("SHIFT " + shift);
    for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(shiftColor);
    dash.setRowHeight(r, 32);
    r++;

    // Sub-row 1: Sample info + inspector + record
    dash.getRange(r,1,1,maxCol).merge().setBackground(shiftColor).setFontColor(WHITE).setFontSize(11).setVerticalAlignment("middle");
    dash.getRange(r,1).setFormula(
      `="Sample: "&TEXT(IFERROR(INDEX(FILTER(${DB}!O2:O,${DB}!A2:A=${LAST_REC}),1),"—"),"MM/DD/YYYY")` +
      `&"  "&TEXT(IFERROR(INDEX(FILTER(${DB}!P2:P,${DB}!A2:A=${LAST_REC}),1),"—"),"hh:mm AM/PM")` +
      `&"  |  Inspector: "&IFERROR(INDEX(FILTER(${DB}!F2:F,${DB}!A2:A=${LAST_REC}),1),"—")` +
      `&"  |  Record: "&IF(${LAST_REC}="","—",${LAST_REC})`
    );
    for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(shiftColor);
    dash.setRowHeight(r, 22);
    r++;

    // Sub-row 2: Inspection info
    dash.getRange(r,1,1,maxCol).merge().setBackground(shiftColor).setFontColor(WHITE).setFontSize(10).setVerticalAlignment("middle");
    dash.getRange(r,1).setFormula(
      `="Inspected: "&TEXT(IFERROR(INDEX(FILTER(${DB}!D2:D,${DB}!A2:A=${LAST_REC}),1),"—"),"MM/DD/YYYY")` +
      `&"  "&TEXT(IFERROR(INDEX(FILTER(${DB}!E2:E,${DB}!A2:A=${LAST_REC}),1),"—"),"hh:mm AM/PM")`
    );
    for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(shiftColor);
    dash.setRowHeight(r, 22);
    r++;

    // Table headers
    for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(DARK);
    tableHeader(r, ["Line","Mold","Cavity","Tests","Pass","Fail","Pass %","Pass / Fail","Trend (last 8 checks)","","","","",""], 1);
    r++;

    for (let i=0; i<lines.length; i++) {
      const lr = r + i, line = lines[i], bg = i%2===1 ? LIGHT : WHITE;

      dataCell(lr, 1, line, null, bg).setFontWeight("bold");
      dataCell(lr, 2, `=IFERROR(INDEX(FILTER(${DB}!K2:K,${DB}!A2:A=${LAST_REC},${DB}!I2:I*1=${line}),1),"")`, null, bg);
      dataCell(lr, 3, `=IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(${DB}!Q2:Q,${DB}!A2:A=${LAST_REC},${DB}!I2:I*1=${line},${DB}!Q2:Q<>""))),"")`, null, bg);
      dataCell(lr, 4, `=IFERROR(COUNTIFS(${DB}!A:A,${LAST_REC},${DB}!I:I,${line},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>"),0)`, "#,##0", bg);
      dataCell(lr, 5, `=IFERROR(COUNTIFS(${DB}!A:A,${LAST_REC},${DB}!I:I,${line},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass"),0)`, "#,##0", bg);
      dataCell(lr, 6, `=IFERROR(COUNTIFS(${DB}!A:A,${LAST_REC},${DB}!I:I,${line},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail"),0)`, "#,##0", bg).setFontColor("#9C0006");
      dataCell(lr, 7, `=IF(D${lr}=0,"",IFERROR(E${lr}/D${lr},""))`, "0.0%", bg).setFontWeight("bold");

      // Stacked bar
      dash.getRange(lr,8).setFormula(`=IF(D${lr}=0,"",SPARKLINE({E${lr},F${lr}},{"charttype","bar";"color1","#27AE60";"color2","#E74C3C"}))`)
        .setBackground(bg).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);

      // Trend sparkline
      dash.getRange(lr,9,1,6).merge()
        .setFormula(`=IFERROR(SPARKLINE(ARRAYFORMULA(IFERROR(COUNTIFS(${DB}!A:A,SORT(UNIQUE(FILTER(${DB}!A2:A,${DB}!G2:G*1=${shift},${DB}!I2:I*1=${line},${DB}!R2:R="Dimensional",${DB}!Y2:Y<>"",${DB}!A2:A<>"")),1,FALSE),${DB}!I:I,${line},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass")/COUNTIFS(${DB}!A:A,SORT(UNIQUE(FILTER(${DB}!A2:A,${DB}!G2:G*1=${shift},${DB}!I2:I*1=${line},${DB}!R2:R="Dimensional",${DB}!Y2:Y<>"",${DB}!A2:A<>"")),1,FALSE),${DB}!I:I,${line},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>"),0)),{"charttype","line";"color","${shiftColor}";"linewidth",2;"ymin",0;"ymax",1}),"")`)
        .setBackground(bg).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    }

    r += lines.length + 1;
    for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(DARK);
    r++;
  }

  // ==================== QC DATA LOOKUP ====================
  sectionHeader(r, "QC DATA LOOKUP — Last 30 days", maxCol);
  r++;

  for (let c=1;c<=maxCol;c++) dash.getRange(r,c).setBackground(MED);

  dash.getRange(r,1).setValue("SHIFT").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r,2).setValue("All").setBackground("#0D2B2B").setFontColor(WHITE).setFontSize(10)
    .setBorder(true,true,true,true,false,false,"#A8C4C4",SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All","1","2","3"],true).setAllowInvalid(false).build());

  dash.getRange(r,4).setValue("MOLD").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r,5).setValue("All").setBackground("#0D2B2B").setFontColor(WHITE).setFontSize(10)
    .setBorder(true,true,true,true,false,false,"#A8C4C4",SpreadsheetApp.BorderStyle.SOLID);
  const dbLast = db.getLastRow();
  const moldData = db.getRange(2,11,dbLast-1,1).getValues().flat();
  const uniqueMolds = [...new Set(moldData.map(m=>String(m).trim()).filter(m=>m))].sort();
  dash.getRange(r,5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["All",...uniqueMolds],true).setAllowInvalid(false).build());

  dash.getRange(r,7).setValue("CHARACTERISTIC").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r,8).setValue("All").setBackground("#0D2B2B").setFontColor(WHITE).setFontSize(10)
    .setBorder(true,true,true,true,false,false,"#A8C4C4",SpreadsheetApp.BorderStyle.SOLID);
  const allChars = ["All","Weight","Bottom Thickness","SW Thickness 1 - Top","SW Thickness 2 - Bottom",
    "SW Thickness 3 - Side Right","SW Thickness 4 - Side Left","Height","Wall Thickness Variance",
    "Top Diameter","Bottom Diameter","End Pin Diameter","Chime A Min","Chime A Max","Chime B Min","Chime B Max",
    "ΔL","ΔA","ΔB","ΔE*ab"];
  dash.getRange(r,8).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(allChars,true).setAllowInvalid(false).build());

  const FSHIFT="$B$"+r, FMOLD="$E$"+r, FCHAR="$H$"+r;

  r++;
  tableHeader(r, ["Inspection Date","Inspected By","Shift","Line","Mold","Cavity","Characteristic","LSL","USL","Actual Value","Status"], 1);
  r++;

  dash.getRange(r,1).setFormula(
    `=IFERROR(QUERY(${DB}!A:AB,"SELECT D, F, G, I, K, Q, T, V, W, X, Y WHERE D >= date '"&TEXT(TODAY()-30,"yyyy-MM-dd")&"' AND D <= date '"&TEXT(TODAY(),"yyyy-MM-dd")&"' AND R = 'Dimensional'"`+
    `&IF(${FSHIFT}="All",""," AND G = "&${FSHIFT}&"")`+
    `&IF(${FMOLD}="All",""," AND K = '"&${FMOLD}&"'")`+
    `&IF(${FCHAR}="All",""," AND T = '"&${FCHAR}&"'")`+
    `&" ORDER BY D DESC, E DESC LABEL D '', F '', G '', I '', K '', Q '', T '', V '', W '', X '', Y ''"),"No data found for selected filters.")`
  ).setFontSize(10).setBorder(true,true,true,true,false,false,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);

  // Column widths
  dash.setColumnWidth(1,110); dash.setColumnWidth(2,140); dash.setColumnWidth(3,60);
  dash.setColumnWidth(4,50); dash.setColumnWidth(5,140); dash.setColumnWidth(6,60);
  dash.setColumnWidth(7,160); dash.setColumnWidth(8,80); dash.setColumnWidth(9,80);
  dash.setColumnWidth(10,90); dash.setColumnWidth(11,70);
  for (let c=12;c<=maxCol;c++) dash.setColumnWidth(c,80);
  dash.setColumnWidth(refCol,1);

  dash.setFrozenRows(3);
  toast_("📋 Shift Performance dashboard built!", "Done");
}


// ================= V3 BACKFILL =================
function backfillV3TestTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName("QC Database");
  if (!db) { SpreadsheetApp.getUi().alert("QC Database tab not found."); return; }

  const lastRow = getLastDataRow_(db);
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data found."); return; }

  const colR = db.getRange(2, 18, lastRow - 1, 1).getValues();
  const colT = db.getRange(2, 20, lastRow - 1, 1).getValues();

  const COLOR_CHARS = ["ΔL","ΔA","ΔB","ΔE*ab"];
  const VISUAL_CHARS = ["Date Code Verification","Visual Conformance Check"];
  const FUNCTIONAL_CHARS = ["Nesting","Cover Fit","Gauge Fit"];
  const DIMENSIONAL_KEYWORDS = [
    "Weight","Bottom Thickness","SW Thickness","Side Wall",
    "Height","Top Diameter","Bottom Diameter","End Pin",
    "Chime","Wall Thickness","Nesting Height",
  ];

  function inferType(charName) {
    const c = String(charName || "").trim();
    if (!c) return "";
    if (COLOR_CHARS.indexOf(c) >= 0) return "Color";
    if (VISUAL_CHARS.indexOf(c) >= 0) return "Visual";
    if (FUNCTIONAL_CHARS.indexOf(c) >= 0) return "Functional";
    for (const kw of DIMENSIONAL_KEYWORDS) {
      if (c.indexOf(kw) >= 0) return "Dimensional";
    }
    return "";
  }

  let fixed = 0, skipped = 0, blank = 0;
  const updates = colR.map((r, i) => {
    const existing = String(r[0] || "").trim();
    if (["Dimensional","Color","Visual","Functional"].indexOf(existing) >= 0) {
      skipped++;
      return [existing];
    }
    const charName = String(colT[i][0] || "").trim();
    const inferred = inferType(charName);
    if (!inferred) { blank++; return [existing]; }
    fixed++;
    return [inferred];
  });

  db.getRange(2, 18, lastRow - 1, 1).setValues(updates);

  SpreadsheetApp.getUi().alert(
    "✅ Backfill complete!\n\n" +
    "Rows fixed: " + fixed + "\n" +
    "Already correct (skipped): " + skipped + "\n" +
    "Could not infer (left blank): " + blank
  );
}


// ================= PROCESS INTELLIGENCE DASHBOARD =================
function buildProcessIntelligenceDashboard() {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DB_SHEET);
  if (!db) { SpreadsheetApp.getUi().alert("QC Database not found."); return; }

  const existing = ss.getSheetByName("📈 Spec & Process Review");
  if (existing) ss.deleteSheet(existing);

  const dash = ss.insertSheet("📈 Spec & Process Review");
  dash.setTabColor("#1D9E75");

  const dbLast = getLastDataRow_(db);
  if (dbLast < 2) { SpreadsheetApp.getUi().alert("No data in QC Database."); return; }

  toast_("Reading database — this may take a moment...", "Building...");

  const rawData = db.getRange(2, 1, dbLast - 1, 25).getValues();

  const moldProd = {};
  const comboData = {};

  for (const row of rawData) {
    const prod = String(row[9] || "").trim();
    const mold = String(row[10] || "").trim();
    const tt = String(row[17] || "").trim();
    const char = String(row[19] || "").trim();
    const lsl = row[21];
    const usl = row[22];
    const actual = row[23];

    if (tt !== "Dimensional" || !mold || !char || typeof actual !== "number") continue;

    const prodN = prod.charAt(0).toUpperCase() + prod.slice(1).toLowerCase();
    moldProd[mold] = prodN;

    const key = mold + "|" + char;
    if (!comboData[key]) comboData[key] = { vals: [], lsl: null, usl: null, prod: prodN, mold, char };
    comboData[key].vals.push(actual);
    if (comboData[key].lsl === null && typeof lsl === "number") comboData[key].lsl = lsl;
    if (comboData[key].usl === null && typeof usl === "number") comboData[key].usl = usl;
  }

  const cpkResults = [];
  for (const d of Object.values(comboData)) {
    if (d.vals.length < 10) continue;
    const n = d.vals.length;
    const avg = d.vals.reduce((a, b) => a + b, 0) / n;
    const std = Math.sqrt(d.vals.reduce((a, b) => a + (b - avg) ** 2, 0) / (n - 1));
    if (std === 0) continue;
    const specRange = (d.usl !== null && d.lsl !== null) ? d.usl - d.lsl : null;
    if (specRange && std > specRange * 5) continue;
    const parts = [];
    if (d.usl !== null) parts.push((d.usl - avg) / (3 * std));
    if (d.lsl !== null) parts.push((avg - d.lsl) / (3 * std));
    if (!parts.length) continue;
    const cpk = Math.min(...parts);
    const cp = (d.usl !== null && d.lsl !== null) ? (d.usl - d.lsl) / (6 * std) : null;
    cpkResults.push({ prod: d.prod, mold: d.mold, char: d.char, n, avg, std, lsl: d.lsl, usl: d.usl, cpk, cp });
  }
  cpkResults.sort((a, b) => a.cpk - b.cpk);

  const prodMolds = {};
  for (const [mold, prod] of Object.entries(moldProd)) {
    if (!prodMolds[prod]) prodMolds[prod] = [];
    if (!prodMolds[prod].includes(mold)) prodMolds[prod].push(mold);
  }
  for (const p in prodMolds) prodMolds[p].sort();

  const allMolds = Object.keys(moldProd).sort();
  const allChars = [...new Set(cpkResults.map(r => r.char))].sort();
  const failMap = {};
  for (const d of cpkResults) failMap[`${d.mold}|${d.char}`] = d.cpk;

  const DARK = "#1A3A3A", MED = "#244747", COPPER = "#E67E22", WHITE = "#FFFFFF", LIGHT = "#F5F5F5";
  const G_BG = "#d4edda", G_FG = "#0f6e56";
  const Y_BG = "#fff3cd", Y_FG = "#856404";
  const R_BG = "#f8d7da", R_FG = "#721c24";
  const maxCol = 14;

  const cpkBg = c => c >= 1.33 ? G_BG : c >= 1.0 ? Y_BG : R_BG;
  const cpkFg = c => c >= 1.33 ? G_FG : c >= 1.0 ? Y_FG : R_FG;
  const cpkLabel = c => c >= 1.33 ? "Capable" : c >= 1.0 ? "Watch" : "Action";
  const bR = rng => rng.setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  function sectionHdr(r, text) {
    const rng = dash.getRange(r, 1, 1, maxCol);
    rng.merge().setValue(text).setBackground(MED).setFontColor(WHITE)
      .setFontSize(12).setFontWeight("bold").setVerticalAlignment("middle");
    for (let c = 1; c <= maxCol; c++) dash.getRange(r, c).setBackground(MED);
  }

  function tblHdr(r, headers, sc) {
    headers.forEach((h, i) => {
      const cell = dash.getRange(r, sc + i);
      cell.setValue(h).setBackground(COPPER).setFontColor(WHITE)
        .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
      bR(cell);
    });
  }

  function writeCell(r, c, val, fmt, bg, fg, bold) {
    const cell = dash.getRange(r, c);
    typeof val === "string" && val.startsWith("=") ? cell.setFormula(val) : cell.setValue(val);
    cell.setFontSize(10).setHorizontalAlignment("center").setBackground(bg || WHITE).setFontColor(fg || "#000000");
    if (bold) cell.setFontWeight("bold");
    if (fmt) cell.setNumberFormat(fmt);
    bR(cell);
    return cell;
  }

  function writeLabel(r, c, text, bg, fg) {
    const cell = dash.getRange(r, c);
    cell.setValue(text).setFontSize(10).setBackground(bg || WHITE).setFontColor(fg || "#000000");
    bR(cell);
    return cell;
  }

  // Spec effective date
  let specEffDate = new Date(2026, 0, 1);
  const storedEff = PropertiesService.getScriptProperties().getProperty("SPEC_EFF_DATE");
  if (storedEff) { const d = new Date(storedEff); if (!isNaN(d.getTime())) specEffDate = d; }

  // Recompute Cpk from spec effective date onward
  const cpkResultsEff = [];
  const comboDataEff = {};
  for (const row of rawData) {
    const mold = String(row[10]||"").trim(), tt = String(row[17]||"").trim();
    const char = String(row[19]||"").trim(), lsl = row[21], usl = row[22];
    const actual = row[23], date = row[3];
    if (tt!=="Dimensional"||!mold||!char||typeof actual!=="number") continue;
    if (!(date instanceof Date)||date<specEffDate) continue;
    const key = mold+"|"+char;
    if (!comboDataEff[key]) comboDataEff[key]={vals:[],lsl:null,usl:null};
    comboDataEff[key].vals.push(actual);
    if (comboDataEff[key].lsl===null&&typeof lsl==="number") comboDataEff[key].lsl=lsl;
    if (comboDataEff[key].usl===null&&typeof usl==="number") comboDataEff[key].usl=usl;
  }
  for (const [key, d] of Object.entries(comboDataEff)) {
    if (d.vals.length<5) continue;
    const kp=key.split("|"), mold2=kp[0], char2=kp.slice(1).join("|");
    const n2=d.vals.length, avg2=d.vals.reduce((a,b)=>a+b,0)/n2;
    const std2=Math.sqrt(d.vals.reduce((a,b)=>a+(b-avg2)**2,0)/(n2-1));
    if (std2===0) continue;
    const sr=(d.usl!==null&&d.lsl!==null)?d.usl-d.lsl:null;
    if (sr&&std2>sr*5) continue;
    const pts=[];
    if (d.usl!==null) pts.push((d.usl-avg2)/(3*std2));
    if (d.lsl!==null) pts.push((avg2-d.lsl)/(3*std2));
    if (!pts.length) continue;
    const cpk2=Math.min(...pts), cp2=(d.usl!==null&&d.lsl!==null)?(d.usl-d.lsl)/(6*std2):null;
    cpkResultsEff.push({prod:moldProd[mold2]||"",mold:mold2,char:char2,n:n2,avg:avg2,std:std2,lsl:d.lsl,usl:d.usl,cpk:cpk2,cp:cp2});
  }
  cpkResultsEff.sort((a,b)=>a.cpk-b.cpk);
  const cpkDisplay=cpkResultsEff.length>0?cpkResultsEff:cpkResults;
  const nAction=cpkDisplay.filter(r=>r.cpk<1.0).length;
  const nWatch=cpkDisplay.filter(r=>r.cpk>=1.0&&r.cpk<1.33).length;
  const nOk=cpkDisplay.filter(r=>r.cpk>=1.33).length;
  for (const d of cpkResultsEff) failMap[`${d.mold}|${d.char}`]=d.cpk;
  const SPEC_EFF_REF="$B$4";

  // Title
  let r = 1;
  dash.setRowHeight(1, 44);
  dash.getRange(1, 1, 1, maxCol).merge().setValue("📈  SPEC & PROCESS REVIEW")
    .setBackground(DARK).setFontColor(WHITE).setFontSize(18).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(2, 24);
  dash.getRange(2, 1, 1, maxCol).merge()
    .setValue("Spec Alignment  ·  Process Capability (Cpk)  ·  SPC Trend Charts  —  Container Supply Co.")
    .setBackground(DARK).setFontColor(COPPER).setFontSize(13).setFontWeight("bold").setVerticalAlignment("middle");

  // Filter bar row 3
  r = 3;
  dash.setRowHeight(3, 32);
  for (let c = 1; c <= maxCol; c++) dash.getRange(r, c).setBackground(MED);
  dash.getRange(r, 1).setValue("START DATE").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r, 2).setValue(new Date(2026, 0, 1)).setBackground("#0D2B2B").setFontColor(WHITE)
    .setFontSize(10).setNumberFormat("mm/dd/yyyy")
    .setBorder(true, true, true, true, false, false, "#A8C4C4", SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 4).setValue("END DATE").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r, 5).setValue(new Date()).setBackground("#0D2B2B").setFontColor(WHITE)
    .setFontSize(10).setNumberFormat("mm/dd/yyyy")
    .setBorder(true, true, true, true, false, false, "#A8C4C4", SpreadsheetApp.BorderStyle.SOLID);

  // Row 4: spec effective date
  r = 4;
  dash.setRowHeight(4, 32);
  for (let c = 1; c <= maxCol; c++) dash.getRange(r, c).setBackground(MED);
  dash.getRange(r, 1).setValue("SPEC DATE").setFontColor(COPPER).setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r, 2).setValue(specEffDate).setBackground("#0D2B2B").setFontColor(COPPER)
    .setFontSize(10).setFontWeight("bold").setNumberFormat("mm/dd/yyyy")
    .setBorder(true, true, true, true, false, false, COPPER, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 4, 1, maxCol-3).merge()
    .setValue("Cpk + heatmap use data from this date forward  |  Use QC Tools > Set Spec Effective Date, then rebuild to refresh Cpk")
    .setFontColor("#A8C4C4").setFontSize(9).setFontStyle("italic").setBackground(MED);

  // KPI strip
  r = 5;
  dash.setRowHeight(5, 52); dash.setRowHeight(6, 20);
  [
    [1, nOk, "Capable  —  Cpk ≥ 1.33", G_BG, G_FG],
    [5, nWatch, "Watch list  —  1.00 to 1.33", Y_BG, Y_FG],
    [9, nAction, "Action required  —  Cpk < 1.00", R_BG, R_FG],
  ].forEach(([sc, val, lbl, bg, fg]) => {
    dash.getRange(r, sc, 1, 3).merge().setValue(val).setFontSize(28).setFontWeight("bold")
      .setFontColor(fg).setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBackground(bg).setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r + 1, sc, 1, 3).merge().setValue(lbl).setFontSize(9).setFontColor(fg)
      .setHorizontalAlignment("center").setBackground(bg);
  });

  // Cpk Table
  r = 8;
  sectionHdr(r, "PROCESS CAPABILITY SUMMARY — sorted by Cpk, worst first");
  r++;
  tblHdr(r, ["Product","Mold","Characteristic","n","Avg","Std Dev","LSL","USL","Cp","Cpk","Status"], 1);
  r++;

  for (let i = 0; i < cpkDisplay.length; i++) {
    const d = cpkDisplay[i];
    const bg = i % 2 === 1 ? LIGHT : WHITE;
    const fmt = d.avg > 10 ? "0.00" : "0.00000";
    writeLabel(r, 1, d.prod, bg);
    writeLabel(r, 2, d.mold, bg);
    writeLabel(r, 3, d.char, bg);
    writeCell(r, 4, d.n, "#,##0", bg);
    writeCell(r, 5, d.avg, fmt, bg);
    writeCell(r, 6, d.std, "0.00000", bg);
    writeCell(r, 7, d.lsl !== null ? d.lsl : "—", fmt, bg, "#0C447C");
    writeCell(r, 8, d.usl !== null ? d.usl : "—", fmt, bg, "#0C447C");
    writeCell(r, 9, d.cp !== null ? d.cp : "—", "0.000", bg);
    writeCell(r, 10, d.cpk, "0.000", cpkBg(d.cpk), cpkFg(d.cpk), true);
    writeCell(r, 11, cpkLabel(d.cpk), null, cpkBg(d.cpk), cpkFg(d.cpk), true);
    r++;
  }

  // Heatmap
  r += 2;
  sectionHdr(r, "FAILURE HEATMAP — Cpk by mold and characteristic");
  r++;

  const HM_CHARS = ["Weight","Bottom Thickness","SW Thickness 1 - Top","SW Thickness 2 - Bottom",
    "SW Thickness 3 - Side Right","SW Thickness 4 - Side Left","Height","Chime A Min","Chime A Max","Chime B Min","Chime B Max"];
  const HM_SHORT = {
    "Weight":"Wt","Bottom Thickness":"BotTh","SW Thickness 1 - Top":"SW1",
    "SW Thickness 2 - Bottom":"SW2","SW Thickness 3 - Side Right":"SW3",
    "SW Thickness 4 - Side Left":"SW4","Height":"Ht",
    "Chime A Min":"ChA-","Chime A Max":"ChA+","Chime B Min":"ChB-","Chime B Max":"ChB+",
  };

  for (const prod of ["Pail","Tub","Cover","Handle"]) {
    const molds = prodMolds[prod];
    if (!molds || molds.length === 0) continue;
    dash.getRange(r, 1, 1, HM_CHARS.length + 1).merge().setValue(prod.toUpperCase())
      .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold");
    for (let c = 1; c <= HM_CHARS.length + 1; c++) dash.getRange(r, c).setBackground(COPPER);
    r++;
    dash.getRange(r, 1).setValue("Mold").setBackground("#244747").setFontColor(WHITE)
      .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
    bR(dash.getRange(r, 1));
    HM_CHARS.forEach((ch, ci) => {
      const cell = dash.getRange(r, ci + 2);
      cell.setValue(HM_SHORT[ch] || ch).setBackground("#244747").setFontColor("#A8C4C4")
        .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
      bR(cell);
    });
    r++;
    molds.forEach((mold, mi) => {
      const rowBg = mi % 2 === 1 ? LIGHT : WHITE;
      writeLabel(r, 1, mold, rowBg);
      HM_CHARS.forEach((ch, ci) => {
        const cpk = failMap[`${mold}|${ch}`];
        const cell = dash.getRange(r, ci + 2);
        if (cpk !== undefined) {
          cell.setValue(cpk.toFixed(2)).setBackground(cpkBg(cpk)).setFontColor(cpkFg(cpk))
            .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
        } else {
          cell.setValue("—").setBackground("#EEEEEE").setFontColor("#AAAAAA")
            .setFontSize(9).setHorizontalAlignment("center");
        }
        bR(cell);
      });
      r++;
    });
    r++;
  }

  // SPC Trend Section
  r++;
  sectionHdr(r, "SPC TREND REVIEW — select a mold and characteristic");
  r++;

  for (let c = 1; c <= maxCol; c++) dash.getRange(r, c).setBackground(MED);
  dash.getRange(r, 1).setValue("MOLD →").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r, 2, 1, 4).merge().setValue(allMolds[0] || "")
    .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, WHITE, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(allMolds, true).setAllowInvalid(false).build()
  );
  const moldRef = `$B$${r}`;

  dash.getRange(r, 7).setValue("CHARACTERISTIC →").setFontColor("#A8C4C4").setFontSize(10).setFontWeight("bold").setBackground(MED);
  dash.getRange(r, 8, 1, 4).merge().setValue(allChars[0] || "")
    .setBackground(COPPER).setFontColor(WHITE).setFontSize(11).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, WHITE, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 8).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(allChars, true).setAllowInvalid(false).build()
  );
  const charRef = `$H$${r}`;
  dash.setRowHeight(r, 36);
  r++;

  // Stats strip
  const statsRow = r;
  const DB = `'QC Database'`;
  const statDefs = [
    ["n",       `=IFERROR(COUNTIFS(${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!X:X,"<>",${DB}!D:D,">="&${SPEC_EFF_REF}),"—")`, "#,##0"],
    ["Avg",     `=IFERROR(AVERAGEIFS(${DB}!X:X,${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!D:D,">="&${SPEC_EFF_REF}),"—")`, "0.0000"],
    ["Std Dev", `=IFERROR(STDEV(FILTER(${DB}!X:X,${DB}!K:K=${moldRef},${DB}!T:T=${charRef},${DB}!R:R="Dimensional",${DB}!D:D>=${SPEC_EFF_REF})),"—")`, "0.00000"],
    ["UCL +3σ", `=IFERROR(B${statsRow+1}+3*C${statsRow+1},"—")`, "0.0000"],
    ["LCL -3σ", `=IFERROR(B${statsRow+1}-3*C${statsRow+1},"—")`, "0.0000"],
    ["LSL",     `=IFERROR(INDEX(FILTER(${DB}!V:V,${DB}!K:K=${moldRef},${DB}!T:T=${charRef},${DB}!D:D>=${SPEC_EFF_REF},${DB}!V:V<>""),1),"—")`, "0.0000"],
    ["USL",     `=IFERROR(INDEX(FILTER(${DB}!W:W,${DB}!K:K=${moldRef},${DB}!T:T=${charRef},${DB}!D:D>=${SPEC_EFF_REF},${DB}!W:W<>""),1),"—")`, "0.0000"],
    ["Cpk",     `=IFERROR(MIN((G${statsRow+1}-B${statsRow+1})/(3*C${statsRow+1}),(B${statsRow+1}-F${statsRow+1})/(3*C${statsRow+1})),"—")`, "0.000"],
  ];

  statDefs.forEach(([lbl, formula, fmt], i) => {
    const lhCell = dash.getRange(r, i + 1);
    lhCell.setValue(lbl).setBackground("#244747").setFontColor("#A8C4C4")
      .setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
    bR(lhCell);
    const valCell = dash.getRange(r + 1, i + 1);
    valCell.setFormula(formula).setFontSize(12).setFontWeight("bold")
      .setHorizontalAlignment("center").setBackground(i === 7 ? "#EFF4FA" : "#F8F8F8")
      .setFontColor(i === 7 ? "#1A3A3A" : "#000000").setNumberFormat(fmt);
    bR(valCell);
  });
  r += 2;

  // Monthly trend
  dash.getRange(r, 1, 1, maxCol).merge()
    .setValue("Monthly trend — updates live when you change mold or characteristic")
    .setBackground("#1A3A3A").setFontColor(COPPER).setFontSize(10).setFontWeight("bold").setVerticalAlignment("middle");
  r++;

  const MT_HDR = ["Month", "n", "Avg", "Min", "Max", "Pass %", "Fails"];
  tblHdr(r, MT_HDR, 1);
  r++;

  const tblStart = r;
  const maxMoRows = 60;
  const stagCol = 16;

  // Staging formula
  dash.getRange(tblStart, stagCol).setFormula(
    `=IFERROR(SORT(UNIQUE(FILTER({${DB}!AD2:AD,${DB}!AC2:AC},${DB}!K2:K=${moldRef},${DB}!T2:T=${charRef},${DB}!R2:R="Dimensional",${DB}!X2:X<>"")),1,FALSE,2,FALSE),"")`
  );
  dash.hideColumns(stagCol, 2);

  const yrL = "P", mthL = "Q";

  const f1=[], f2=[], f3=[], f4=[], f5=[], f6=[], f7=[];
  for (let i = 0; i < maxMoRows; i++) {
    const vr = tblStart + i;
    const isEmpty = `NOT(ISNUMBER(${yrL}${vr}))`;
    const mRef = `${mthL}${vr}`;
    const yRef = `${yrL}${vr}`;

    f1.push([`=IF(${isEmpty},"",TEXT(DATE(${yRef},${mRef},1),"MMM YYYY"))`]);
    f2.push([`=IF(${isEmpty},"",COUNTIFS(${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!X:X,"<>",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}))`]);
    f3.push([`=IF(${isEmpty},"",IFERROR(AVERAGEIFS(${DB}!X:X,${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}),""))`]);
    f4.push([`=IF(${isEmpty},"",IFERROR(MINIFS(${DB}!X:X,${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}),""))`]);
    f5.push([`=IF(${isEmpty},"",IFERROR(MAXIFS(${DB}!X:X,${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}),""))`]);
    f6.push([`=IF(${isEmpty},"",IFERROR(COUNTIFS(${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Pass",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef})/COUNTIFS(${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"<>",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}),""))`]);
    f7.push([`=IF(${isEmpty},"",COUNTIFS(${DB}!K:K,${moldRef},${DB}!T:T,${charRef},${DB}!R:R,"Dimensional",${DB}!Y:Y,"Fail",${DB}!AD:AD,${yRef},${DB}!AC:AC,${mRef}))`]);
  }

  dash.getRange(tblStart, 1, maxMoRows, 1).setFormulas(f1);
  dash.getRange(tblStart, 2, maxMoRows, 1).setFormulas(f2);
  dash.getRange(tblStart, 3, maxMoRows, 1).setFormulas(f3);
  dash.getRange(tblStart, 4, maxMoRows, 1).setFormulas(f4);
  dash.getRange(tblStart, 5, maxMoRows, 1).setFormulas(f5);
  dash.getRange(tblStart, 6, maxMoRows, 1).setFormulas(f6);
  dash.getRange(tblStart, 7, maxMoRows, 1).setFormulas(f7);

  dash.getRange(tblStart, 3, maxMoRows, 1).setNumberFormat("0.0000");
  dash.getRange(tblStart, 4, maxMoRows, 1).setNumberFormat("0.0000");
  dash.getRange(tblStart, 5, maxMoRows, 1).setNumberFormat("0.0000");
  dash.getRange(tblStart, 6, maxMoRows, 1).setNumberFormat("0%");
  dash.getRange(tblStart, 7, maxMoRows, 1).setNumberFormat("#,##0");

  dash.getRange(tblStart, 1, maxMoRows, 7)
    .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  for (let i = 0; i < maxMoRows; i++) {
    dash.getRange(tblStart + i, 1, 1, 7).setBackground(i % 2 === 1 ? LIGHT : WHITE);
    dash.setRowHeight(tblStart + i, 20);
  }

  // Column widths
  dash.setColumnWidth(1, 70);
  dash.setColumnWidth(2, 165);
  dash.setColumnWidth(3, 195);
  dash.setColumnWidth(4, 50);
  dash.setColumnWidth(5, 70);
  dash.setColumnWidth(6, 75);
  dash.setColumnWidth(7, 65);
  dash.setColumnWidth(8, 65);
  dash.setColumnWidth(9, 60);
  dash.setColumnWidth(10, 65);
  dash.setColumnWidth(11, 80);
  for (let c = 12; c <= maxCol; c++) dash.setColumnWidth(c, 65);

  dash.setFrozenRows(3);
  toast_("📈 Spec & Process Review built!", "Done");
}


// ================= SPEC EFFECTIVE DATE HELPER =================
function setSpecEffectiveDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Set Spec Effective Date",
    "Enter the date your current specs became valid (MM/DD/YYYY).\n\n" +
    "Cpk table uses only data from this date forward.\nRebuild the dashboard after changing.",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const dateStr = resp.getResponseText().trim();
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) { ui.alert("Invalid date. Use MM/DD/YYYY."); return; }

  PropertiesService.getScriptProperties().setProperty("SPEC_EFF_DATE", date.toISOString());

  const formatted = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
  const dash = SpreadsheetApp.getActive().getSheetByName("📈 Spec & Process Review");
  if (dash) {
    dash.getRange("B4").setValue(date).setNumberFormat("mm/dd/yyyy");
    ui.alert("Spec date set to " + formatted + "\n\nRebuild Spec & Process Review to refresh Cpk.");
  } else {
    ui.alert("Saved: " + formatted + "\n\nBuild Spec & Process Review to apply it.");
  }
}

