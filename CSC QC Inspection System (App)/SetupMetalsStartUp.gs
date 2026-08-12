/*************************************************************
 * ONE-TIME: creates "Start-Up Verification Items List- Metals" and
 * "Start-Up Verification Log- Metals", seeding the items list with a draft checklist
 * built from the Metals start-up process as described: Job Readiness (line/material
 * staging, mirrored from the Plastics form's Job Readiness category), then Slitter
 * (body blank dimensions + artwork/litho inspection), then Double Seamer (tear-down
 * confirmation + end lot data).
 *
 * A few rows are educated-guess placeholders, called out below and in their Notes column —
 * review and adjust in the sheet once you see the rendered form:
 *   - Body Blank Weight, Squareness / Diagonal (Length and Width were the only two confirmed)
 *   - The exact wording of the artwork/litho defect checks (out-of-register, color, smearing)
 *   - Job Readiness wording adapted from Plastics' resin/press/mold language to Metals'
 *     coil/line/tooling equivalents — confirm the exact raw-material terminology
 *
 * Run once from the Apps Script editor (select populateMetalsStartUpItems, click Run).
 *************************************************************/
// [Verification Item, Value Type, Unit, Notes, Category] — mirrors the Plastics
// Start-Up Verification Items List's "Job Readiness" category.
const METALS_JOB_READINESS_ITEMS = [
  ['Previous job fully cleared from line — no remnant coil, stock, or tags', 'Yes/No', '', 'PLACEHOLDER wording — confirm Metals raw-material terminology', 'Job Readiness'],
  ['Guards, shields, and covers are in place and secured', 'Yes/No', '', '', 'Job Readiness'],
  ['Product contact surfaces are clean — no residue, oil, or debris', 'Yes/No', '', '', 'Job Readiness'],
  ['Correct tooling / dies installed and confirmed for this part number', 'Yes/No', '', '', 'Job Readiness'],
  ['Set-up technician has verified and signed off on machine setup', 'Yes/No', '', '', 'Job Readiness'],
  ['Correct raw material (coil / sheet stock) on hand per the production order', 'Yes/No', '', 'PLACEHOLDER wording — confirm Metals raw-material terminology', 'Job Readiness'],
  ['No mixed or unidentified materials present at the line', 'Yes/No', '', '', 'Job Readiness'],
  ['The line is ready, staged, and ready for production', 'Yes/No', '', '', 'Job Readiness'],
];

function populateMetalsStartUpItems() {
  const ss = getDb_();
  ensureSheetWithHeaders_(ss, SU_LOG_SHEET_NAME_METALS, METALS_SU_LOG_HEADERS);

  const sheet = ensureSheetWithHeaders_(ss, SU_ITEMS_SHEET_NAME_METALS, ['Verification Item', 'Value Type', 'Unit', 'Notes', 'Category']);
  if (sheet.getLastRow() > 1) {
    SpreadsheetApp.getActive().toast('Start-Up Verification Items List- Metals already has data — leaving it as-is.');
    return;
  }

  // [Verification Item, Value Type, Unit, Notes, Category]
  const ITEMS = [
    ...METALS_JOB_READINESS_ITEMS,

    ['Body Blank Length', 'Number', 'in', '', 'Slitter — Body Blank Dimensions'],
    ['Body Blank Width', 'Number', 'in', '', 'Slitter — Body Blank Dimensions'],
    ['Body Blank Weight', 'Number', 'lb', 'PLACEHOLDER — confirm this is an actual start-up check and its units', 'Slitter — Body Blank Dimensions'],
    ['Squareness / Diagonal', 'Number', 'in', 'PLACEHOLDER — confirm this is an actual start-up check and its units', 'Slitter — Body Blank Dimensions'],

    ['Artwork Correct Per Order', 'Yes/No', '', '', 'Slitter — Artwork / Litho Inspection'],
    ['Free of Out-of-Register Defects', 'Yes/No', '', '', 'Slitter — Artwork / Litho Inspection'],
    ['Correct Color Match', 'Yes/No', '', '', 'Slitter — Artwork / Litho Inspection'],
    ['Free of Smearing', 'Yes/No', '', '', 'Slitter — Artwork / Litho Inspection'],

    ['Double Seamer Tear-Down Completed & Passing Specs', 'Yes/No', '', 'Detailed measurements are logged in separate seamer inspection software — this just confirms it was done and passed', 'Double Seamer — Tear-Down & Ends'],
    ['Correct Ends Verified Per Order', 'Yes/No', '', '', 'Double Seamer — Tear-Down & Ends'],
    ['End Item No.', 'Text', '', 'For lot traceability', 'Double Seamer — Tear-Down & Ends'],
    ['End Description', 'Text', '', '', 'Double Seamer — Tear-Down & Ends'],
    ['End Lot Date', 'Date', '', 'Also serves as date of manufacture for recall traceability', 'Double Seamer — Tear-Down & Ends'],

    ['Was there a deviation?', 'Yes/No', '', '', ''],
    ['Deviation Description / Notes', 'Text', '', '', ''],
  ];

  sheet.getRange(2, 1, ITEMS.length, 5).setValues(ITEMS);
  SpreadsheetApp.getActive().toast('Added ' + ITEMS.length + ' draft checklist items to Start-Up Verification Items List- Metals. Review/adjust in the sheet, then test the form.');
}

/**
 * ONE-TIME: adds the Job Readiness category to an already-populated
 * Start-Up Verification Items List- Metals (populateMetalsStartUpItems only seeds an
 * empty sheet, so this is the follow-up for a sheet that already has the Slitter /
 * Double Seamer rows). Inserts the rows at the top so Job Readiness renders first on
 * the form, same as it does on the Plastics checklist.
 * Run once from the Apps Script editor (select addMetalsJobReadinessItems, click Run).
 */
function addMetalsJobReadinessItems() {
  const sheet = getDb_().getSheetByName(SU_ITEMS_SHEET_NAME_METALS);
  if (!sheet) throw new Error('"' + SU_ITEMS_SHEET_NAME_METALS + '" sheet not found — run populateMetalsStartUpItems first.');

  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues().map(r => r[0]);
  if (existing.indexOf(METALS_JOB_READINESS_ITEMS[0][0]) >= 0) {
    SpreadsheetApp.getActive().toast('Job Readiness items are already in Start-Up Verification Items List- Metals — leaving it as-is.');
    return;
  }

  sheet.insertRowsBefore(2, METALS_JOB_READINESS_ITEMS.length);
  sheet.getRange(2, 1, METALS_JOB_READINESS_ITEMS.length, 5).setValues(METALS_JOB_READINESS_ITEMS);
  SpreadsheetApp.getActive().toast('Added ' + METALS_JOB_READINESS_ITEMS.length + ' Job Readiness items to Start-Up Verification Items List- Metals.');
}
