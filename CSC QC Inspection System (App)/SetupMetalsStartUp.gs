/*************************************************************
 * ONE-TIME: creates "Start-Up Verification Items List- Metals" and
 * "Start-Up Verification Log- Metals", seeding the items list with a draft checklist
 * built from the Metals start-up process as described: Slitter (body blank dimensions +
 * artwork/litho inspection) then Double Seamer (tear-down confirmation + end lot data).
 *
 * A few rows are educated-guess placeholders, called out below and in their Notes column —
 * review and adjust in the sheet once you see the rendered form:
 *   - Body Blank Weight, Squareness / Diagonal (Length and Width were the only two confirmed)
 *   - The exact wording of the artwork/litho defect checks (out-of-register, color, smearing)
 *
 * Run once from the Apps Script editor (select populateMetalsStartUpItems, click Run).
 *************************************************************/
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
