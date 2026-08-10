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
    columns: ['QC Record #', 'Inspection Date', 'Mold', 'Product Type', 'Characteristic Name', 'Actual Value', 'Status', 'Inspected By', 'Line #'],
  },
  dropfreeze: {
    sheetName: DROPFREEZE_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Result', dateField: 'TestDate',
    columns: ['RecordKey', 'Run ID', 'Mold ID', 'Test Name', 'Cavity', 'TestDate', 'TestedBy', 'Result'],
  },
  startup: {
    sheetName: SU_LOG_SHEET_NAME, moldField: 'Mold ID', statusField: 'Status', dateField: 'Verification Date',
    columns: ['Verification Record #', 'Run ID', 'Mold ID', 'Verification Item', 'Actual Value', 'Status', 'Verification Date', 'QC Tech Name'],
  },
};

const REVIEW_ROW_CAP_ = 500;

function sanitizeForClient_(v) { return v instanceof Date ? dateToStr_(v) : v; }

function toDateSafe_(v) {
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * filters: {moldId, status, dateFrom, dateTo, search} — all optional, all substring/range matches.
 * Returns { columns, rows, totalMatched, capped }. Rows are sorted newest-first by the module's
 * date field and capped at REVIEW_ROW_CAP_ — totalMatched/capped tell the client if more exist.
 */
function getTestDataReviewRows(module, filters) {
  const cfg = REVIEW_MODULES_[module];
  if (!cfg) throw new Error('Unknown review module: ' + module);
  const sheet = getDb_().getSheetByName(cfg.sheetName);
  if (!sheet) return { columns: cfg.columns, rows: [], totalMatched: 0, capped: false };

  filters = filters || {};
  const moldFilter = String(filters.moldId || '').trim().toLowerCase();
  const statusFilter = String(filters.status || '').trim().toLowerCase();
  const searchFilter = String(filters.search || '').trim().toLowerCase();
  const dateFrom = filters.dateFrom ? toDateSafe_(filters.dateFrom) : null;
  const dateTo = filters.dateTo ? toDateSafe_(filters.dateTo) : null;

  let rows = readSheetObjects_(sheet).filter(r => {
    if (moldFilter && String(r[cfg.moldField] || '').toLowerCase().indexOf(moldFilter) < 0) return false;
    if (statusFilter && String(r[cfg.statusField] || '').toLowerCase().indexOf(statusFilter) < 0) return false;
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

  const outRows = rows.map(r => cfg.columns.map(c => sanitizeForClient_(r[c])));
  return { columns: cfg.columns, rows: outRows, totalMatched: totalMatched, capped: capped };
}
