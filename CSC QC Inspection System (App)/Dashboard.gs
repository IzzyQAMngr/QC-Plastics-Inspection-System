/*************************************************************
 * PLASTICS DASHBOARD — one card per currently Active Run, showing
 * that run's In-Process and Drop Freeze pass rates, most recent Resin
 * Lot, and Qualified status. Computed live off the log sheets on every
 * load (no separate dashboard tabs are written) — Resin Lot in
 * particular is read straight off the In-Process log rather than kept
 * in some imported/synced copy, same as everything else here.
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

  // Drop Freeze rows are matched by Run ID, not a timestamp, and this log used to be nowhere
  // near In-Process's size — a full read here wasn't the slow part. It's grown enough since
  // (see listOpenDropFreezeRecords_ in DropFreeze.gs, fixed for the same reason) that it's
  // worth bounding the same way: a sample can never be logged before the Run it belongs to was
  // created, so it's always safe to skip rows older than every currently-Active run's start.
  const dropFreezeSheet = getDb_().getSheetByName(DROPFREEZE_LOG_SHEET_NAME);
  const dropFreezeRows = dropFreezeSheet ? readSheetObjectsSince_(dropFreezeSheet, 'Created', earliestCreatedAt) : [];

  // Start-Up Verification: drafts (in-progress, not yet submitted) and deviation status
  // (submitted with a deviation, pending Manager approval, or approved and awaiting PFA
  // sign-off) — both keyed by Run ID so buildLineCard_ can just look its Run up.
  const drafts = getDraftedRunIds_('Plastics');
  const suStatusByRun = getStartUpStatusByRun_('Plastics');

  const cards = runs.map(run => buildLineCard_(run, inProcessRows, dropFreezeRows, drafts[run.runId], suStatusByRun[run.runId]));
  cards.sort((a, b) => {
    const na = parseFloat(a.line), nb = parseFloat(b.line);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : String(a.line).localeCompare(String(b.line));
  });

  return { cards: cards, generatedAt: dateToStr_(new Date()) };
}

/** One pass over the Start-Up Verification log, grouped into records (by Verification Record
 *  #) and then reduced to the MOST RECENTLY SAVED record per Run ID — mirrors
 *  getPendingDeviations_'s grouping but keyed by Run instead of listed by record, and kept for
 *  every deviation regardless of whether it's still pending, since an already-qualified Run
 *  still needs to know its qualification came through an approved deviation (for the
 *  Dashboard's small flag) rather than a clean pass. Runs with no deviation at all are simply
 *  absent from the returned map. */
function getStartUpStatusByRun_(department) {
  const sheet = getDb_().getSheetByName(getSuLogSheetName_(department));
  if (!sheet) return {};
  const rows = readSheetObjects_(sheet);
  const byRecord = {};
  rows.forEach(r => {
    const id = r['Verification Record #'];
    if (!id) return;
    (byRecord[id] = byRecord[id] || []).push(r);
  });

  const latestByRun = {};
  Object.keys(byRecord).forEach(recordId => {
    const group = byRecord[recordId];
    const runId = String(group[0]['Run ID'] || '').trim();
    if (!runId) return;
    const raw = group[0]['Timestamp saved'];
    const ts = raw instanceof Date ? raw : new Date(raw);
    const prev = latestByRun[runId];
    if (!prev || (!isNaN(ts.getTime()) && ts > prev.ts)) latestByRun[runId] = { ts: ts, group: group };
  });

  const out = {};
  Object.keys(latestByRun).forEach(runId => {
    const group = latestByRun[runId].group;
    const hasDeviation = group.some(r => r['Verification Item'] === 'Was there a deviation?' && String(r['Actual Value']).trim() === 'Yes');
    if (!hasDeviation) return;
    const approved = group.some(r => r['Verification Item'] === 'Deviation Approved by' && String(r['Actual Value']).trim());
    const signedOff = group.some(r => r['Verification Item'] === 'PFA Signed off by' && String(r['Actual Value']).trim());
    const descRow = group.find(r => r['Verification Item'] === 'Deviation Description / Notes');
    out[runId] = {
      status: signedOff ? 'Qualified' : (approved ? 'Awaiting PFA Sign-off' : 'Awaiting Approval'),
      deviationDescription: descRow ? String(descRow['Actual Value'] || '').trim() : '',
    };
  });
  return out;
}

function buildLineCard_(run, inProcessRows, dropFreezeRows, draftInfo, suStatus) {
  const createdAt = run.createdAt ? new Date(run.createdAt) : null;

  // Line #/Mold/Run ID cells can come back from Sheets as either a number or a string
  // depending on how that row was written, so both sides must be normalized to strings
  // before comparing — comparing a bare (possibly numeric) run.line/run.moldId against an
  // already-stringified sheet value silently matched nothing.
  let ipPass = 0, ipFail = 0;
  let latestResinLot = '', latestResinLotAt = null;
  const charStats = {}; // Characteristic Name -> {pass, fail} — which dimensions are driving a Line's fails
  inProcessRows.forEach(r => {
    if (String(r['Line #'] || '').trim() !== String(run.line || '').trim()) return;
    if (String(r['Mold'] || '').trim() !== String(run.moldId || '').trim()) return;
    const t = r['Timestamp Saved'] instanceof Date ? r['Timestamp Saved'] : new Date(r['Timestamp Saved']);
    if (createdAt && !isNaN(createdAt.getTime()) && !isNaN(t.getTime()) && t < createdAt) return;

    // Resin Lot is captured per sample on every In-Process row regardless of Test Type (Visual,
    // Color, Dimensional, Functional rows from the same save all carry the same "LOT of Resin"),
    // so it's tracked here off the whole matched set, not just the Dimensional rows the pass/fail
    // tally below is scoped to — whichever row has the latest Timestamp Saved wins.
    const resinLot = String(r['LOT of Resin'] || '').trim();
    if (resinLot && !isNaN(t.getTime()) && (!latestResinLotAt || t > latestResinLotAt)) {
      latestResinLot = resinLot;
      latestResinLotAt = t;
    }

    if (r['Test Type'] !== 'Dimensional') return;
    const status = String(r.Status || '').trim();
    if (status === 'Pass') ipPass++;
    else if (status === 'Fail') ipFail++;
    else return;
    const charName = String(r['Characteristic Name'] || '').trim();
    if (!charName) return;
    if (!charStats[charName]) charStats[charName] = { pass: 0, fail: 0 };
    charStats[charName][status === 'Pass' ? 'pass' : 'fail']++;
  });

  // Only the characteristics actually failing, worst contributor first — this is what tells a
  // QC Manager glancing at the board WHY a Line's In-Process rate is low, not just that it is.
  // `rate` is each characteristic's share of the LINE'S OVERALL total (ipPass + ipFail), not of
  // that characteristic's own sample count — so every failingChars[].rate plus the gauge's own
  // pass rate always adds up to 100%, matching what's shown above it (e.g. 90% pass + 10% fail,
  // with a single failing characteristic accounting for the whole 10%).
  const ipTotal = ipPass + ipFail;
  const failingChars = Object.keys(charStats)
    .map(name => {
      const s = charStats[name];
      return { name, fail: s.fail, total: ipTotal, rate: ipTotal ? Math.round((s.fail / ipTotal) * 1000) / 10 : 0 };
    })
    .filter(c => c.fail > 0)
    .sort((a, b) => b.fail - a.fail)
    .slice(0, 5);

  let dfPass = 0, dfFail = 0;
  dropFreezeRows.forEach(r => {
    if (String(r['Run ID'] || '').trim() !== String(run.runId || '').trim()) return;
    if (String(r.Status || '').trim().toUpperCase() !== 'COMPLETE') return;
    if (isFailValue_(r.Result)) dfFail++;
    else if (String(r.Result || '').toUpperCase().indexOf('PASS') >= 0) dfPass++;
  });

  return {
    runId: run.runId, line: run.line, shift: run.shift, createdAt: run.createdAt,
    moldId: run.moldId, moldDescription: run.moldDescription, color: run.color,
    item: run.item, itemDescription: run.itemDescription,
    runQty: run.runQty, resinLot: latestResinLot, resinLotAsOf: dateToStr_(latestResinLotAt),
    qualified: run.qualified === 'Yes',
    hasDraft: !!draftInfo, draftSavedBy: draftInfo ? draftInfo.savedBy : '',
    deviationStatus: suStatus ? suStatus.status : '', deviationDescription: suStatus ? suStatus.deviationDescription : '',
    inProcess: { pass: ipPass, fail: ipFail, total: ipPass + ipFail, failingChars: failingChars },
    dropFreeze: { pass: dfPass, fail: dfFail, total: dfPass + dfFail },
  };
}
