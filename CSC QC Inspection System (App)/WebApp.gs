/*************************************************************
 * WEB APP ENTRY POINT
 *************************************************************/

// Shown under the logo in the sidebar so it's obvious at a glance which build is live —
// bump this alongside every `clasp deploy` to the production deployment ID (see
// reference_deployment_details memory), matching the @N version number clasp reports.
const APP_VERSION = 'v125';

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

/** renamePrefix lets the SAME partial be included more than once on one page (every view's
 *  HTML+script is concatenated into one document, all at once, regardless of which .view is
 *  currently visible) without its element ids or global JS names colliding — e.g.
 *  DropFreezeResultsForm.html uses a "dfrf" prefix throughout; a second inclusion elsewhere on
 *  the page passes a different prefix here to get its own independent copy. */
function include(filename, renamePrefix) {
  const html = HtmlService.createHtmlOutputFromFile(filename).getContent();
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
