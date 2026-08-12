/*************************************************************
 * PLASTICS DASHBOARD — one card per currently Active Run, showing
 * that run's In-Process and Drop Freeze pass rates plus its
 * Qualified status. Computed live off the log sheets on every load
 * (no separate dashboard tabs are written).
 *
 * In-Process has no Run ID column (only Line # + Mold), so its rows
 * are matched to a Run by Line # + Mold + Timestamp Saved on/after
 * the Run's Created At — a close approximation, not a hard link.
 * Drop Freeze rows carry Run ID directly, so that match is exact.
 *************************************************************/

function getPlasticsLineDashboardData() {
  const runs = getActiveRuns_('Plastics');

  const inProcessSheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  const inProcessRows = inProcessSheet ? readSheetObjects_(inProcessSheet) : [];

  const dropFreezeSheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  const dropFreezeRows = dropFreezeSheet ? readSheetObjects_(dropFreezeSheet) : [];

  const cards = runs.map(run => buildLineCard_(run, inProcessRows, dropFreezeRows));
  cards.sort((a, b) => {
    const na = parseFloat(a.line), nb = parseFloat(b.line);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : String(a.line).localeCompare(String(b.line));
  });

  return { cards: cards, generatedAt: dateToStr_(new Date()) };
}

function buildLineCard_(run, inProcessRows, dropFreezeRows) {
  const createdAt = run.createdAt ? new Date(run.createdAt) : null;

  let ipPass = 0, ipFail = 0;
  inProcessRows.forEach(r => {
    if (String(r['Line #'] || '').trim() !== run.line) return;
    if (String(r['Mold'] || '').trim() !== run.moldId) return;
    if (createdAt && !isNaN(createdAt.getTime())) {
      const t = r['Timestamp Saved'] instanceof Date ? r['Timestamp Saved'] : new Date(r['Timestamp Saved']);
      if (!isNaN(t.getTime()) && t < createdAt) return;
    }
    if (r['Test Type'] !== 'Dimensional') return;
    const status = String(r.Status || '').trim();
    if (status === 'Pass') ipPass++;
    else if (status === 'Fail') ipFail++;
  });

  let dfPass = 0, dfFail = 0;
  dropFreezeRows.forEach(r => {
    if (String(r['Run ID'] || '').trim() !== run.runId) return;
    if (String(r.Status || '').trim().toUpperCase() !== 'COMPLETE') return;
    if (isFailValue_(r.Result)) dfFail++;
    else if (String(r.Result || '').toUpperCase().indexOf('PASS') >= 0) dfPass++;
  });

  return {
    runId: run.runId, line: run.line, shift: run.shift, createdAt: run.createdAt,
    moldId: run.moldId, moldDescription: run.moldDescription,
    item: run.item, itemDescription: run.itemDescription, customerName: run.customerName,
    qualified: run.qualified === 'Yes',
    inProcess: { pass: ipPass, fail: ipFail, total: ipPass + ipFail },
    dropFreeze: { pass: dfPass, fail: dfFail, total: dfPass + dfFail },
  };
}
