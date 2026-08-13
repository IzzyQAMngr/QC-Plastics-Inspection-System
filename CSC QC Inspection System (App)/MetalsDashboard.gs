/*************************************************************
 * METALS DASHBOARD — one card per currently Active Metals Run, showing its
 * Qualified status (from Start-Up Verification). The meter slots mirror the
 * Plastics Dashboard's card shape (In-Process / drop-or-functional-test pass
 * rates) but stay at "No data yet" until the Metals Inspection module (see
 * MetalsView.html) is actually built and has a log sheet to read from.
 *************************************************************/

function getMetalsLineDashboardData() {
  const runs = getActiveRuns_('Metals');

  const cards = runs.map(run => ({
    runId: run.runId, line: run.line, shift: run.shift, createdAt: run.createdAt,
    sizeId: run.sizeId, canDescription: run.canDescription, productType: run.productType,
    item: run.item, itemDescription: run.itemDescription, customerName: run.customerName,
    qualified: run.qualified === 'Yes',
    metalsInProcess: { pass: 0, fail: 0, total: 0 },
    dropFunctional: { pass: 0, fail: 0, total: 0 },
  }));

  cards.sort((a, b) => {
    const na = parseFloat(a.line), nb = parseFloat(b.line);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : String(a.line).localeCompare(String(b.line));
  });

  return { cards: cards, generatedAt: dateToStr_(new Date()) };
}
