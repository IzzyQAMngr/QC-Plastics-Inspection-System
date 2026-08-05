/*************************************************************
 * HOME DASHBOARD — live KPIs across product families.
 * Computed directly from the log sheets (no separate dashboard
 * tabs are written) — deeper drill-down dashboards are Phase 2/3.
 *************************************************************/

const WEEK_MS = 7 * 24 * 60 * 60 * 1000;

function getHomeDashboardData() {
  return {
    dropFreeze: getDropFreezeKpis_(),
    inProcess: getInProcessKpis_(),
    generatedAt: new Date().toISOString(),
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
  const byToolCode = {};

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
    const tc = String(r.ToolCode || '').trim();
    if (tc) {
      byToolCode[tc] = byToolCode[tc] || { toolCode: tc, total: 0, fail: 0 };
      byToolCode[tc].total++;
      if (isFailValue_(r.Result)) byToolCode[tc].fail++;
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
    byToolCode: Object.values(byToolCode).sort((a, b) => b.fail - a.fail).slice(0, 8),
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
        byMold[mold] = byMold[mold] || { mold: mold, fail: 0 };
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
