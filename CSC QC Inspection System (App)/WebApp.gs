/*************************************************************
 * WEB APP ENTRY POINT
 *************************************************************/
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CSC QC Inspection System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1200)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSC QC Inspection System');
}

// Update this if the web app is ever deployed under a different deployment ID.
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxEh-YeIBHgXj2hHktCXnibKLlic35euwkv9W6hs8HfAdGdUSQWhlSnUb1EwyQuRjqEOg/exec';

function showWebAppLink() {
  SpreadsheetApp.getUi().alert('Open the deployed app', WEB_APP_URL, SpreadsheetApp.getUi().ButtonSet.OK);
}
