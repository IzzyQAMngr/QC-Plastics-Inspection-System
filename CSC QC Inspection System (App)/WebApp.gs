/*************************************************************
 * WEB APP ENTRY POINT
 *************************************************************/

// Shown under the logo in the sidebar so it's obvious at a glance which build is live —
// bump this alongside every `clasp deploy` to the production deployment ID (see
// reference_deployment_details memory), matching the @N version number clasp reports.
const APP_VERSION = 'v156';

function doGet(e) {
  const params = (e && e.parameter) || {};
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.initialParamsJson = JSON.stringify({ view: params.view || '', record: params.record || '', dept: params.dept || '' });
  tmpl.appVersion = APP_VERSION;
  return tmpl.evaluate()
    .setTitle('CSC QC Inspection System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Polled from the client (see Index.html's update-banner logic) so a tab left open across a
 *  deploy — e.g. the Run Dashboard, meant to sit on a wall monitor for days — finds out its
 *  page code is stale without anyone having to manually refresh it. */
function getAppVersion() { return APP_VERSION; }

/** renamePrefix lets the SAME partial be included more than once on one page (every view's
 *  HTML+script is concatenated into one document, all at once, regardless of which .view is
 *  currently visible) without its element ids or global JS names colliding — e.g.
 *  DropFreezeResultsForm.html uses a "dfrf" prefix throughout; a second inclusion elsewhere on
 *  the page passes a different prefix here to get its own independent copy.
 *  Must use createTemplateFromFile(...).evaluate(), NOT createHtmlOutputFromFile(...).getContent()
 *  — the latter returns the file's raw text with no scriptlet evaluation, so any <?!= include(...) ?>
 *  tag INSIDE an included partial (e.g. DropFreezeView.html including DropFreezeResultsForm.html)
 *  was left as literal unevaluated text in the page instead of being replaced (found 2026-09-10:
 *  this silently broke Open Samples' and Drop Freeze Test's results form since 2026-09-04, look
 *  like a hang because the literal tag meant osfrfInit/dfrfInit were never defined, throwing when
 *  called and aborting the render right after the data fetch had already logged success). */
function include(filename, renamePrefix) {
  const html = HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  return renamePrefix ? html.split('dfrf').join(renamePrefix) : html;
}

// ================= SPREADSHEET MENU (preview the app without deploying) =================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSC QC Inspection System')
    .addItem('🔍 Preview App', 'showAppDialog')
    .addSeparator()
    .addItem('🌐 Open Deployed Web App', 'showWebAppLink')
    .addToUi();
}

function showAppDialog() {
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.initialParamsJson = '{}';
  tmpl.appVersion = APP_VERSION;
  const html = tmpl.evaluate()
    .setWidth(1200)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSC QC Inspection System');
}

// Update this if the web app is ever deployed under a different deployment ID.
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxEh-YeIBHgXj2hHktCXnibKLlic35euwkv9W6hs8HfAdGdUSQWhlSnUb1EwyQuRjqEOg/exec';

function showWebAppLink() {
  SpreadsheetApp.getUi().alert('Open the deployed app', WEB_APP_URL, SpreadsheetApp.getUi().ButtonSet.OK);
}
