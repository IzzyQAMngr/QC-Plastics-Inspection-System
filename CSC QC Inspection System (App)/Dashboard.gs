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
  if (runs.length === 0) return { cards: [], generatedAt: dateToStr_(new Date()) };

  // Only rows saved on/after the earliest Active run's start can possibly match one of these
  // runs — on an 80k+-row log, reading the whole thing just to filter it down in JS is the
  // reason this call used to take 27-33s every load. See readSheetObjectsSince_.
  const earliestCreatedAt = new Date(Math.min.apply(null, runs.map(r => new Date(r.createdAt).getTime())));

  const inProcessSheet = getDb_().getSheetByName(INPROCESS_LOG_SHEET_NAME);
  const inProcessRows = inProcessSheet ? readSheetObjectsSince_(inProcessSheet, 'Timestamp Saved', earliestCreatedAt) : [];

  // Drop Freeze rows are matched by Run ID, not a timestamp, and this log is nowhere near
  // In-Process's size — a full read here isn't the slow part.
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

  // Line #/Mold/Run ID cells can come back from Sheets as either a number or a string
  // depending on how that row was written, so both sides must be normalized to strings
  // before comparing — comparing a bare (possibly numeric) run.line/run.moldId against an
  // already-stringified sheet value silently matched nothing.
  let ipPass = 0, ipFail = 0;
  inProcessRows.forEach(r => {
    if (String(r['Line #'] || '').trim() !== String(run.line || '').trim()) return;
    if (String(r['Mold'] || '').trim() !== String(run.moldId || '').trim()) return;
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
    if (String(r['Run ID'] || '').trim() !== String(run.runId || '').trim()) return;
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
