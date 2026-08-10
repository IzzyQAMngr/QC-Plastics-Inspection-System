/*************************************************************
 * ADD RUN — creates/stops the Runs that In-Process and Drop Freeze both pull from.
 * A Run replaces the old ad-hoc Batch: it's set up once here, stays selectable while
 * Active (across shifts, across days), and carries the Qualified flag set by a passing
 * Start-Up Verification submission.
 *************************************************************/

/** Called by AddRunView.html on load. */
function getAddRunFormData() {
  return {
    molds: getAllMoldsList_(),        // [{moldId, description, productType}]
    itemList: getItemList_(),         // [{itemNo, description}]
    shifts: getShiftList_(),
    createdByOptions: getInspectorList_(), // QC Technician Name roster — the only staff list on file today
    activeRuns: getActiveRuns_(),
  };
}

/** Public wrapper — fields: {shift, line, productType, resinLot, moldId, moldDescription,
 *  item, itemDescription, customerName, runQty, createdBy}. */
function addRun(fields) { return createRun_(fields); }

function endRun(runId) { return stopRun_(runId); }

function confirmRunsToday(confirmedBy) { return confirmTodaysRuns_(confirmedBy); }
