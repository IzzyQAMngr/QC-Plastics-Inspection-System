/*************************************************************
 * ADD RUN — METALS — mirrors AddRun.gs's Plastics behavior, writing into the separate
 * "Runs - Metals" sheet. Can Size replaces Mold, Material Lot replaces Resin Lot.
 *************************************************************/

/** Called by MetalsAddRunView.html on load. */
function getAddMetalsRunFormData() {
  return {
    sizeCans: getAllSizeCansList_(),        // [{sizeId, productType}]
    itemList: getItemList_('Metals'),       // [{itemNo, description}]
    lines: getLinesForDepartment_('Metals'),
    shifts: getShiftList_('Metals'),
    createdByOptions: getInspectorList_('Metals'),
    activeRuns: getActiveRuns_('Metals'),
  };
}

/** Public wrapper — fields: {shift, line, productType, materialLot, sizeId, canDescription,
 *  item, itemDescription, customerName, runQty, createdBy}. */
function addMetalsRun(fields) { return createRun_(fields, 'Metals'); }

function endMetalsRun(runId) { return stopRun_(runId, 'Metals'); }

function confirmMetalsRunsToday(confirmedBy) { return confirmTodaysRuns_(confirmedBy, 'Metals'); }
