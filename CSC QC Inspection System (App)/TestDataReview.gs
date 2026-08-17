/*************************************************************
 * TEST DATA REVIEW — browse/filter saved inspection records across
 * In-Process, Drop Freeze, and Start-Up Verification. Each module's log
 * has a different shape, so this reads generically off a per-module config
 * (which sheet, which columns to show, which fields back the Mold/Status/
 * Date filters) rather than three near-duplicate functions.
 *************************************************************/

const REVIEW_MODULES_ = {
  inprocess: {
    sheetName: INPROCESS_LOG_SHEET_NAME, moldField: 'Mold', statusField: 'Status', dateField: 'Timestamp Saved',
    charField: 'Characteristic Name', charLabel: 'Characteristic',
    itemField: 'Item No.', colorField: 'Color', valueField: 'Actual Value', unitField: 'Unit',
    columns: ['QC Record #', 'Inspection Date', 'Mold', 'Item No.', 'Color', 'Product Type', 'Characteristic Name', 'Actual Value', 'Status', 'Inspected By', 'Line #'],
  },
  dropfreeze: {
    sheetName: DROPFREEZE_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Result', dateField: 'TestDate',
    charField: 'Test Name', charLabel: 'Test',
    itemField: 'Item No', colorField: null, valueField: null, unitField: null,
    columns: ['RecordKey', 'Run ID', 'Mold ID', 'Item No', 'Test Name', 'Cavity', 'TestDate', 'TestedBy', 'Result'],
  },
  startup: {
    sheetName: SU_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Status', dateField: 'Verification Date',
    charField: 'Verification Item', charLabel: 'Verification Item',
    itemField: 'Item', colorField: null, valueField: 'Actual Value', unitField: 'Unit',
    columns: ['Verification Record #', 'Run ID', 'Mold ID', 'Item', 'Verification Item', 'Actual Value', 'Status', 'Verification Date', 'QC Tech Name'],
  },
};

const REVIEW_ROW_CAP_ = 500;

/**
 * readSheetObjects_ (Shared.gs) reads every column and builds a full-width object per row — fine
 * for small sheets, but In-Process's log has grown to 80k+ rows across 35 columns, and this review
 * only ever touches ~13 of them. Reading all 35 for every keystroke-triggered filter is most of
 * why this page felt slow. This reads only the requested columns, grouped into contiguous runs so
 * a handful of getValues() calls cover them instead of one call per column.
 */
function readReviewLogRows_(sheet, neededHeaders) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const allHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const wanted = new Set(neededHeaders);
  const wantedCols = [];
  allHeaders.forEach((h, i) => { if (h && wanted.has(h)) wantedCols.push({ idx: i, header: h }); });
  if (!wantedCols.length) return [];

  const numRows = lastRow - 1;
  const rows = new Array(numRows);
  for (let r = 0; r < numRows; r++) rows[r] = {};

  let runStart = 0;
  for (let i = 1; i <= wantedCols.length; i++) {
    if (i < wantedCols.length && wantedCols[i].idx === wantedCols[i - 1].idx + 1) continue;
    const run = wantedCols.slice(runStart, i);
    const values = sheet.getRange(2, run[0].idx + 1, numRows, run.length).getValues();
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < run.length; c++) rows[r][run[c].header] = values[r][c];
    }
    runStart = i;
  }
  return rows;
}

/** Every field the review ever reads off a module's log — the display columns plus whichever
 *  filter/stat fields aren't already in that list (e.g. In-Process's date filter runs off
 *  Timestamp Saved, not the Inspection Date column shown in the table). */
function reviewNeededHeaders_(cfg) {
  const extra = [cfg.moldField, cfg.itemField, cfg.colorField, cfg.charField, cfg.statusField, cfg.dateField, cfg.valueField, cfg.unitField];
  return Array.from(new Set(cfg.columns.concat(extra.filter(Boolean))));
}

/** Every displayed date column here (Inspection Date, TestDate, Verification Date) is a plain
 *  calendar date with no meaningful time-of-day, so format as yyyy-MM-dd instead of a raw ISO
 *  timestamp — dateToStr_'s full ISO string is still what drives date-range filtering/sorting,
 *  this only affects what's shown in the table. */
function sanitizeForClient_(v, tz) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v; }

function toDateSafe_(v) {
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Fixed Status option list, read from the "Inspection Status List" column in Settings (add it there
 * the same way Shift/QC Technician Name/etc. are set up — a header cell plus the values below it).
 * Deliberately NOT derived by scanning the log sheets — same values regardless of module.
 */
function getReviewStatusOptions_() { return settingsColumnBelow_('Inspection Status List', 30); }
function getReviewStatusOptions() { return getReviewStatusOptions_(); }

/**
 * Distinct Characteristic/Item Number/Color values for the module's own log data — all
 * high-cardinality (dimensional points, verification items, item numbers, colors) and vary by
 * mold, so there's no sensible fixed list for any of them. Lazy-loaded client-side (first time
 * Filters is expanded for a module), not on every switch, and in one scan of the sheet.
 */
function getReviewFilterOptions(module) {
  const cfg = REVIEW_MODULES_[module];
  if (!cfg) throw new Error('Unknown review module: ' + module);
  const sheet = getDb_().getSheetByName(cfg.sheetName);
  const optionFields = Array.from(new Set([cfg.charField, cfg.itemField, cfg.colorField].filter(Boolean)));
  const rows = sheet ? readReviewLogRows_(sheet, optionFields) : [];
  const characteristics = new Set();
  const itemNumbers = new Set();
  const colors = new Set();
  rows.forEach(r => {
    const c = String(r[cfg.charField] || '').trim(); if (c) characteristics.add(c);
    if (cfg.itemField) { const it = String(r[cfg.itemField] || '').trim(); if (it) itemNumbers.add(it); }
    if (cfg.colorField) { const co = String(r[cfg.colorField] || '').trim(); if (co) colors.add(co); }
  });
  return {
    charLabel: cfg.charLabel,
    characteristics: Array.from(characteristics).sort((a, b) => a.localeCompare(b)),
    itemNumbers: Array.from(itemNumbers).sort((a, b) => a.localeCompare(b)),
    colors: Array.from(colors).sort((a, b) => a.localeCompare(b)),
    hasColor: !!cfg.colorField,
  };
}

/**
 * Summary stats over ALL matched rows (before the display cap) — so it reflects the true filtered
 * set even when results are capped. A count-by-status breakdown works for every module, including
 * Drop Freeze, whose "value" is a Pass/Fail/Inconclusive result rather than a number. Average/min/max
 * only apply where cfg.valueField is configured, and only over the rows that actually parse as a
 * number — Start-Up mixes numeric verification items with text/dropdown/sign-off ones.
 */
function computeReviewStats_(rows, cfg) {
  const byStatus = {};
  const nums = [];
  let unit = '';
  rows.forEach(r => {
    const s = String(r[cfg.statusField] || '').trim() || '(blank)';
    byStatus[s] = (byStatus[s] || 0) + 1;
    if (cfg.valueField) {
      const n = parseFloat(r[cfg.valueField]);
      if (!isNaN(n) && isFinite(n)) {
        nums.push(n);
        if (!unit && cfg.unitField && r[cfg.unitField]) unit = String(r[cfg.unitField]).trim();
      }
    }
  });
  const numeric = nums.length ? {
    count: nums.length,
    avg: nums.reduce((a, b) => a + b, 0) / nums.length,
    min: Math.min.apply(null, nums),
    max: Math.max.apply(null, nums),
    unit: unit,
  } : null;
  return { total: rows.length, byStatus: byStatus, numeric: numeric };
}

/**
 * filters: {moldId, itemNo, color, characteristic, search, dateFrom, dateTo, status} — all optional.
 * Mold/Item Number/Color/Characteristic/search are typed, substring matches. Status is a single exact
 * match off the fixed Settings list. Returns { columns, rows, totalMatched, capped, stats }. Rows are
 * sorted newest-first by the module's date field and capped at REVIEW_ROW_CAP_ — totalMatched/capped
 * tell the client if more exist; stats is computed over every matched row, not just the capped page.
 */
function getTestDataReviewRows(module, filters) {
  const cfg = REVIEW_MODULES_[module];
  if (!cfg) throw new Error('Unknown review module: ' + module);
  const sheet = getDb_().getSheetByName(cfg.sheetName);
  if (!sheet) return { columns: cfg.columns, rows: [], totalMatched: 0, capped: false, stats: { total: 0, byStatus: {}, numeric: null } };

  filters = filters || {};
  const moldFilter = String(filters.moldId || '').trim().toLowerCase();
  const itemFilter = String(filters.itemNo || '').trim().toLowerCase();
  const colorFilter = String(filters.color || '').trim().toLowerCase();
  const charFilter = String(filters.characteristic || '').trim().toLowerCase();
  const statusFilter = String(filters.status || '').trim().toLowerCase();
  const searchFilter = String(filters.search || '').trim().toLowerCase();
  const dateFrom = filters.dateFrom ? toDateSafe_(filters.dateFrom) : null;
  const dateTo = filters.dateTo ? toDateSafe_(filters.dateTo) : null;

  let rows = readReviewLogRows_(sheet, reviewNeededHeaders_(cfg)).filter(r => {
    if (moldFilter && String(r[cfg.moldField] || '').toLowerCase().indexOf(moldFilter) < 0) return false;
    if (itemFilter && cfg.itemField && String(r[cfg.itemField] || '').toLowerCase().indexOf(itemFilter) < 0) return false;
    if (colorFilter && cfg.colorField && String(r[cfg.colorField] || '').toLowerCase().indexOf(colorFilter) < 0) return false;
    if (charFilter && String(r[cfg.charField] || '').toLowerCase().indexOf(charFilter) < 0) return false;
    if (statusFilter && String(r[cfg.statusField] || '').trim().toLowerCase() !== statusFilter) return false;
    if (dateFrom || dateTo) {
      const d = toDateSafe_(r[cfg.dateField]);
      if (!d) return false;
      if (dateFrom && d < dateFrom) return false;
      if (dateTo && d > dateTo) return false;
    }
    if (searchFilter) {
      const hit = Object.keys(r).some(k => String(r[k] || '').toLowerCase().indexOf(searchFilter) >= 0);
      if (!hit) return false;
    }
    return true;
  });

  rows.sort((a, b) => {
    const da = toDateSafe_(a[cfg.dateField]), db = toDateSafe_(b[cfg.dateField]);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });

  const totalMatched = rows.length;
  const stats = computeReviewStats_(rows, cfg);
  const capped = totalMatched > REVIEW_ROW_CAP_;
  rows = rows.slice(0, REVIEW_ROW_CAP_);

  const tz = getDb_().getSpreadsheetTimeZone();
  const outRows = rows.map(r => cfg.columns.map(c => sanitizeForClient_(r[c], tz)));
  return { columns: cfg.columns, rows: outRows, totalMatched: totalMatched, capped: capped, stats: stats };
}
