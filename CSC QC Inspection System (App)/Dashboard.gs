/*************************************************************
 * HOME DASHBOARD — live KPIs + bar charts across product families.
 * Computed directly from the log sheets (no separate dashboard
 * tabs are written) — deeper drill-down dashboards are Phase 3.
 *************************************************************/

const WEEK_MS = 7 * 24 * 60 * 60 * 1000;

function getHomeDashboardData() {
  return {
    dropFreeze: getDropFreezeKpis_(),
    inProcess: getInProcessKpis_(),
    startUp: getStartUpKpis_(),
    generatedAt: dateToStr_(new Date()),
  };
}

function isRecentDate_(value, withinMs) {
  const d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return false;
  return (new Date().getTime() - d.getTime()) <= withinMs;
}

function getDropFreezeKpis_() {
  const sheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  if (!sheet) return { available: false };
  const rows = readSheetObjects_(sheet);

  let openCount = 0, completeCount = 0, passCount = 0, failCount = 0, failsThisWeek = 0;
  const byMold = {};

  rows.forEach(r => {
    const status = String(r.Status || '').trim().toUpperCase();
    if (status === 'OPEN') openCount++;
    if (status === 'COMPLETE') {
      completeCount++;
      if (isFailValue_(r.Result)) {
        failCount++;
        if (isRecentDate_(r.Updated, WEEK_MS)) failsThisWeek++;
      } else if (String(r.Result || '').toUpperCase().indexOf('PASS') >= 0) {
        passCount++;
      }
    }
    const mold = String(r['Mold ID'] || '').trim();
    if (mold) {
      byMold[mold] = byMold[mold] || { label: mold, total: 0, fail: 0 };
      byMold[mold].total++;
      if (isFailValue_(r.Result)) byMold[mold].fail++;
    }
  });

  const graded = passCount + failCount;
  return {
    available: true,
    totalRecords: rows.length,
    open: openCount,
    complete: completeCount,
    passRate: graded ? Math.round((passCount / graded) * 1000) / 10 : null,
    failsThisWeek: failsThisWeek,
    byMold: Object.values(byMold).sort((a, b) => b.fail - a.fail).slice(0, 8),
  };
}

function getInProcessKpis_() {
  const sheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  if (!sheet) return { available: false };
  const rows = readSheetObjects_(sheet);

  const recordIds = new Set();
  let dimPass = 0, dimFail = 0, failsThisWeek = 0;
  const byMold = {};

  rows.forEach(r => {
    recordIds.add(String(r['QC Record #'] || '').trim());
    const status = String(r.Status || '').trim();
    if (r['Test Type'] === 'Dimensional') {
      if (status === 'Pass') dimPass++;
      if (status === 'Fail') dimFail++;
    }
    if (status === 'Fail') {
      if (isRecentDate_(r['Timestamp Saved'], WEEK_MS)) failsThisWeek++;
      const mold = String(r.Mold || '').trim();
      if (mold) {
        byMold[mold] = byMold[mold] || { label: mold, fail: 0 };
        byMold[mold].fail++;
      }
    }
  });

  const graded = dimPass + dimFail;
  return {
    available: true,
    totalRecords: recordIds.size,
    dimPassRate: graded ? Math.round((dimPass / graded) * 1000) / 10 : null,
    failsThisWeek: failsThisWeek,
    byMold: Object.values(byMold).sort((a, b) => b.fail - a.fail).slice(0, 8),
  };
}

/** Start-Up Verification + Run qualification rollups. */
function getStartUpKpis_() {
  const runsSheet = getDb_().getSheetByName(RUNS_SHEET_NAME);
  const logSheet = getDb_().getSheetByName(SU_LOG_SHEET_NAME);
  if (!runsSheet || !logSheet) return { available: false };

  const runs = readSheetObjects_(runsSheet);
  const qualifiedCount = runs.filter(r => String(r.Qualified || '').trim() === 'Yes').length;
  const runQualRate = runs.length ? Math.round((qualifiedCount / runs.length) * 1000) / 10 : null;

  const logRows = readSheetObjects_(logSheet);
  const byRecord = {};
  logRows.forEach(r => {
    const id = String(r['Verification Record #'] || '').trim();
    if (!id) return;
    (byRecord[id] = byRecord[id] || []).push(r);
  });

  let deviationCount = 0;
  const byMold = {};
  Object.keys(byRecord).forEach(id => {
    const group = byRecord[id];
    const hasDeviation = group.some(r => r['Verification Item'] === 'Was there a deviation?' && String(r['Actual Value']).trim() === 'Yes');
    if (!hasDeviation) return;
    deviationCount++;
    const mold = String(group[0]['Mold ID'] || '').trim();
    if (mold) { byMold[mold] = byMold[mold] || { label: mold, fail: 0 }; byMold[mold].fail++; }
  });

  return {
    available: true,
    totalVerifications: Object.keys(byRecord).length,
    totalRuns: runs.length,
    qualifiedRuns: qualifiedCount,
    runQualRate: runQualRate,
    deviationsReported: deviationCount,
    pendingApprovals: getPendingDeviations_().length,
    byMold: Object.values(byMold).sort((a, b) => b.fail - a.fail).slice(0, 8),
  };
}
