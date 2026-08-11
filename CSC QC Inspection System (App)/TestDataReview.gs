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
    columns: ['QC Record #', 'Inspection Date', 'Mold', 'Product Type', 'Characteristic Name', 'Actual Value', 'Status', 'Inspected By', 'Line #'],
  },
  dropfreeze: {
    sheetName: DROPFREEZE_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Result', dateField: 'TestDate',
    charField: 'Test Name', charLabel: 'Test',
    columns: ['RecordKey', 'Run ID', 'Mold ID', 'Test Name', 'Cavity', 'TestDate', 'TestedBy', 'Result'],
  },
  startup: {
    sheetName: SU_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Status', dateField: 'Verification Date',
    charField: 'Verification Item', charLabel: 'Verification Item',
    columns: ['Verification Record #', 'Run ID', 'Mold ID', 'Verification Item', 'Actual Value', 'Status', 'Verification Date', 'QC Tech Name'],
  },
};

const REVIEW_ROW_CAP_ = 500;

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
 * Distinct Characteristic values for the module's own log data — the one filter still worth
 * scanning for, since it's high-cardinality (dimensional points, verification items, etc.) and
 * varies by mold, so there's no sensible fixed list for it. Lazy-loaded client-side (first time
 * Filters is expanded for a module), not on every switch.
 */
function getReviewFilterOptions(module) {
  const cfg = REVIEW_MODULES_[module];
  if (!cfg) throw new Error('Unknown review module: ' + module);
  const sheet = getDb_().getSheetByName(cfg.sheetName);
  const rows = sheet ? readSheetObjects_(sheet) : [];
  const characteristics = new Set();
  rows.forEach(r => { const c = String(r[cfg.charField] || '').trim(); if (c) characteristics.add(c); });
  return { charLabel: cfg.charLabel, characteristics: Array.from(characteristics).sort((a, b) => a.localeCompare(b)) };
}

/**
 * filters: {moldId, characteristic, search, dateFrom, dateTo, status} — all optional.
 * Mold/Characteristic/search are typed, substring matches. Status is a single exact match off the
 * fixed Settings list. Returns { columns, rows, totalMatched, capped }. Rows are sorted newest-first
 * by the module's date field and capped at REVIEW_ROW_CAP_ — totalMatched/capped tell the client if
 * more exist.
 */
function getTestDataReviewRows(module, filters) {
  const cfg = REVIEW_MODULES_[module];
  if (!cfg) throw new Error('Unknown review module: ' + module);
  const sheet = getDb_().getSheetByName(cfg.sheetName);
  if (!sheet) return { columns: cfg.columns, rows: [], totalMatched: 0, capped: false };

  filters = filters || {};
  const moldFilter = String(filters.moldId || '').trim().toLowerCase();
  const charFilter = String(filters.characteristic || '').trim().toLowerCase();
  const statusFilter = String(filters.status || '').trim().toLowerCase();
  const searchFilter = String(filters.search || '').trim().toLowerCase();
  const dateFrom = filters.dateFrom ? toDateSafe_(filters.dateFrom) : null;
  const dateTo = filters.dateTo ? toDateSafe_(filters.dateTo) : null;

  let rows = readSheetObjects_(sheet).filter(r => {
    if (moldFilter && String(r[cfg.moldField] || '').toLowerCase().indexOf(moldFilter) < 0) return false;
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
  const capped = totalMatched > REVIEW_ROW_CAP_;
  rows = rows.slice(0, REVIEW_ROW_CAP_);

  const tz = getDb_().getSpreadsheetTimeZone();
  const outRows = rows.map(r => cfg.columns.map(c => sanitizeForClient_(r[c], tz)));
  return { columns: cfg.columns, rows: outRows, totalMatched: totalMatched, capped: capped };
}
