/**
 * MENU SETUP
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Trip Tools')
    .addItem('Add Dispatch Details', 'openDispatchDialog')
    .addItem('Calculate Time (50 MPH)', 'GET_TRIP_TIME')
    .addItem('Calculate Distance (Nearest Mile)', 'GET_TRIP_DISTANCE')
    .addToUi();
}

function openDispatchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog')
      .setWidth(500)
      .setHeight(450)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paste Dispatch Details');
}