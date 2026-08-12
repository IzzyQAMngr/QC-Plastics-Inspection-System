/*************************************************************
 * ONE-TIME: populate Line Configuration with Metals equipment, pulled from the
 * Equipment Inventory Register (FRM-030-001). Each "ENTIRE LINE" summary asset
 * row was dropped — only individual equipment is listed. Run once from the
 * Apps Script editor (select populateMetalsLineConfiguration, click Run).
 *************************************************************/
function populateMetalsLineConfiguration() {
  const sheet = getDb_().getSheetByName(LINE_CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('"' + LINE_CONFIG_SHEET_NAME + '" sheet not found — run setupLineConfiguration() first.');

  const existing = readSheetObjects_(sheet);
  if (existing.some(r => String(r['Department'] || '').trim() === 'Metals')) {
    SpreadsheetApp.getActive().toast('Metals rows already exist in Line Configuration — leaving them as-is.');
    return;
  }

  // [Line #, Equipment Code, Equipment Description]
  const METALS_EQUIPMENT = [
    ['2', '21', 'SLITTER'], ['2', '22', 'BODYMAKER/NORDSON'], ['2', '23', 'OVEN/OVEN CONVEY.'],
    ['2', '24', 'FLANGER'], ['2', '25', 'BEADER'], ['2', '26', 'SEAMER'], ['2', '27', 'TESTER'],
    ['2', '28', 'PALLETIZER'], ['2', '29', 'ELEVATORS/RUNWAYS'],

    ['2B', '21B', 'SLITTER'], ['2B', '22B', 'BODYMAKER/NORDSON'], ['2B', '23B', 'OVEN/OVEN CONVEY.'],
    ['2B', '24B', 'FLANGER'], ['2B', '25B', 'BEADER'], ['2B', '26B', 'SEAMER'], ['2B', '27B', 'TESTER'],
    ['2B', '28B', 'PALLETIZER'], ['2B', '29B', 'ELEVATORS/RUNWAYS'],

    ['3', '31', 'SLITTER'], ['3', '32', 'BODYMAKER/NORDSON'], ['3', '33', 'OVEN/OVEN CONVEY.'],
    ['3', '34', 'FLANGER'], ['3', '35', 'BEADER'], ['3', '36', 'SEAMER'], ['3', '37', 'TESTER'],
    ['3', '38', 'PALLETIZER'], ['3', '39', 'ELEVATORS/RUNWAYS'],

    ['4', '41', 'SLITTER #4A'], ['4', '41-1', 'SLITTER #4B'], ['4', '42', 'BODYMAKER/NORDSON'],
    ['4', '43', 'OVEN/OVEN CONVEY.'], ['4', '44', 'FLANGER/BEADER'], ['4', '46', 'SEAMER'],
    ['4', '47', 'TESTER'], ['4', '48', 'PALLETIZER'], ['4', '49', 'ELEVATORS/RUNWAYS'],

    ['5', '51', 'SLITTER #5A'], ['5', '52', 'BODYMAKER/NORDSON'], ['5', '53', 'OVEN/OVEN CONVEY.'],
    ['5', '54', 'FLANGER'], ['5', '55', 'BEADER'], ['5', '56', 'SEAMER'], ['5', '57', 'TESTER'],
    ['5', '58', 'PALLETIZER'], ['5', '59', 'ELEVATORS/RUNWAYS'],

    ['6', '61', 'SLITTER #6A(603X700)'], ['6', '62', 'BODYMAKER/NORDSON'],
    ['6', '62-1', 'BODYMAKER/NORDSON/POWDER COATING'], ['6', '63', 'OVEN/OVEN CONVEYORS'],
    ['6', '63-1', 'OVEN/OVEN CONVEYORS'], ['6', '64', 'FLANGER/BEADER'], ['6', '64-1', 'FLANGER/BEADER/SEAMER'],
    ['6', '66-1', 'SEAMER'], ['6', '67', 'TESTER'], ['6', '68', 'PALLETIZER/WRAPPER'],
    ['6', '68-1', 'DOUBLE HIGH PAL./WRAP.'], ['6', '69', 'ELEVATORS/RUNWAYS'], ['6', '69-1', 'ELEVATORS/RUNWAYS'],

    ['Hand Line', '74', 'SEAMER'],

    ['Slitter', '110', 'ALL'], ['Slitter', '111', 'NO. 1H'], ['Slitter', '112', 'NO. 2H'],
    ['Slitter', '113', 'NO. 3H'], ['Slitter', '114', 'NO. 4H'], ['Slitter', '115', 'NO. 5H'],
    ['Slitter', '116', 'NO. 10A (610X711)'], ['Slitter', '117', 'NO. 7A (603X812)'], ['Slitter', '118', 'NO. 8A'],
    ['Slitter', '119', 'NO. 9A (401/603 ENDS)'], ['Slitter', '120', 'NO. 1 (OLDER 603)'],
    ['Slitter', '122', 'NO. 2 (401)'], ['Slitter', '123', 'NO. 3 (NEWER 603)'],

    ['Coil Line 1', '131', 'UNCOILER'], ['Coil Line 1', '132', 'SLITTER'], ['Coil Line 1', '133', 'STRAIGHTENER'],
    ['Coil Line 1', '134', 'SHEAR'], ['Coil Line 1', '135', 'CONVEYORS'], ['Coil Line 1', '136', 'UPENDER'],
    ['Coil Line 1', '137', 'SERVO ROLL FEEDER'],

    ['Coil Line 2', '126-1', 'P. HOLE CLASS./STACKER'], ['Coil Line 2', '127-1', 'OVER/UNDER CLASS./STKR.'],
    ['Coil Line 2', '128-1', 'CLASSIFIER/STKR. #1'], ['Coil Line 2', '129-1', 'CLASSIFIER/STKR. #2'],
    ['Coil Line 2', '131-1', 'UNCOILER'], ['Coil Line 2', '132-1', 'TRIMMER'], ['Coil Line 2', '133-1', 'STRAIGHTENER'],
    ['Coil Line 2', '134-1', 'SHEAR'], ['Coil Line 2', '135-1', 'CONVEYORS'], ['Coil Line 2', '136-1', 'UPENDER'],
    ['Coil Line 2', '138-1', 'INSP. TABLE'], ['Coil Line 2', '139-1', 'PIN HOLE DET.'],

    ['Press', '150', 'ALL/SUPPLIES'], ['Press', '151', 'NO. 1'], ['Press', '152', 'NO. 2'],
    ['Press', '163', 'NO. 13'], ['Press', '164', 'NO. 14'], ['Press', '165', 'NO. 15'], ['Press', '166', 'NO. 16'],
    ['Press', '167', 'NO. 17'], ['Press', '168', 'NO. 18'], ['Press', '169', 'NO. 19'],
    ['Press', '170', 'TWO WHEEL'], ['Press', '171', 'SEGMENT'], ['Press', '172', 'ALL'],
    ['Press', '173', 'AUTO PRESS NO. 20'], ['Press', '174', 'CEVIS - CAN END VISUAL INSPECTION SYSTEM'],
    ['Press', '200', 'ALL'], ['Press', '202', 'MINI-SHINER DIE'], ['Press', '203', '216 STUD HOLE DIE'],
    ['Press', '204', '212 PLUG DIE'], ['Press', '205', '212 OPENING DIE'], ['Press', '206', '211 AUTO DOUBLE DIE'],
    ['Press', '207', '3" OPENING DIE'], ['Press', '208', '307 SANITARY END DIE'], ['Press', '209', '307 END DIE'],
    ['Press', '210', '307 AUTO DOUBLE DIE'], ['Press', '215', '401 AUTO END DIE'], ['Press', '216', '401 END DIE'],
    ['Press', '217', '401 AUTO DOUBLE DIE'], ['Press', '218', '401 HYDRAULIC END DIE'],
    ['Press', '221', '404 AUTO DOUBLE DIE'], ['Press', '222', '404 PLAIN END DIE--A'],
    ['Press', '223', '404 PLAIN END DIE--B'], ['Press', '224', '404 INK SLIP COVER DIE'],
    ['Press', '225', '404 SLIP COVER DIE'], ['Press', '226', '404 F/F PLUG DIE'],
    ['Press', '227', '408 X 502 RING DIE'], ['Press', '228', '408 S/F PLUG DIE'],
    ['Press', '229', '414 F/F PLUG DIE'], ['Press', '230', '414 END DIE'], ['Press', '231', '414 FOLLOWER PLATE DIE'],
    ['Press', '232', '502 F/F COVER DIE'], ['Press', '233', '502 AUTO END DIE'], ['Press', '234', '502 "A" END DIE'],
    ['Press', '235', '502 SLIP COVER DIE'], ['Press', '236', '502 "B" END DIE'], ['Press', '237', '504 END DIE'],
    ['Press', '238', '504 F/F PLUG DIE'], ['Press', '239', '507 OPENING DIE'], ['Press', '240', '507 S/F PLUG DIE'],
    ['Press', '241', '509 END DIE'], ['Press', '242', '509 F/F PLUG DIE'], ['Press', '243', '515 END DIE'],
    ['Press', '245', '515 SLIP COVER DIE'], ['Press', '246', '6" SLIP COVER DIE'], ['Press', '247', '602 END DIE'],
    ['Press', '248', '602 INK SLIP COVER DIE'], ['Press', '249', '602 CTC END DIE'], ['Press', '250', '6" END DIE'],
    ['Press', '251', '603 AUTO END DIE'], ['Press', '253', '603 S/F RING DIE'], ['Press', '255', '603 END DIE'],
    ['Press', '256', '610 2ND. OPER. D/F PLUG DIE'], ['Press', '257', '610 2ND. OPER. D/F RING DIE'],
    ['Press', '258', '610 1ST. OPER. D/F PLUG DIE'], ['Press', '259', '610 1ST. OPER. D/F RING DIE'],
    ['Press', '260', '610 END DIE'], ['Press', '262', '610 F/F PLUG DIE'],
    ['Press', '263', '610 3RD. OPER. D/F RING DIE'], ['Press', '264', '715 SLIP COVER DIE'],
    ['Press', '265', '715 S/F RING DIE'], ['Press', '266', '715 END DIE'], ['Press', '267', '7 1/4 FRICTION CAP DIE'],
    ['Press', '268', '808 SLIP COVER DIE'], ['Press', '269', '808 END DIE'], ['Press', '270', '9" FRUIT INSERT DIE'],
    ['Press', '271', '10" KENNECOTT END DIE'], ['Press', '272', '10" A END DIE'], ['Press', '273', '10" B END DIE'],
    ['Press', '274', '10" SHALLOW COVER DIE'], ['Press', '275', '10" SLIP COVER DIE'],
    ['Press', '277', 'REIKE JR. DIE'], ['Press', '278', '610 AUTO END DIE'],
  ];

  const rows = METALS_EQUIPMENT.map(r => ({ 'Department': 'Metals', 'Line #': r[0], 'Equipment Code': r[1], 'Equipment Description': r[2] }));
  appendObjectsAsRows_(sheet, rows);
  SpreadsheetApp.getActive().toast('Added ' + rows.length + ' Metals equipment rows to Line Configuration.');
}

/** Diagnostic: shows exactly what the app sees in the Metals Item List tab, so a header/tab
 *  name mismatch can be spotted directly instead of guessed at. Run once, read the popup. */
function debugMetalsItemList() {
  const ss = getDb_();
  const lines = [];
  const sheet = ss.getSheetByName(ALL_ITEMS_LIST_SHEET_NAME_METALS);
  if (!sheet) {
    lines.push('Sheet "' + ALL_ITEMS_LIST_SHEET_NAME_METALS + '" NOT FOUND in this spreadsheet.');
  } else {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const lastRow = sheet.getLastRow();
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    lines.push('Sheet "' + ALL_ITEMS_LIST_SHEET_NAME_METALS + '" found: ' + lastRow + ' rows x ' + lastCol + ' cols.');
    lines.push('Row 1 (header): [' + headerRow.map(h => '"' + h + '"').join(', ') + ']');
    if (lastRow > 1) {
      const sample = sheet.getRange(2, 1, Math.min(3, lastRow - 1), lastCol).getValues();
      sample.forEach((r, i) => lines.push('Row ' + (i + 2) + ': [' + r.map(v => '"' + v + '"').join(', ') + ']'));
    }
    const pos = findHeaderRowAndCol_(sheet, 'Item No.', 3);
    lines.push(pos ? 'Code found "Item No." at row ' + pos.row + ', col ' + pos.col + '.' : 'Code did NOT find a header cell reading exactly "Item No." in the first 3 rows.');
  }
  const msg = lines.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Metals Item List Diagnostic', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* headless */ }
}
