/**
 * Atlas — Dynamic Dashboard (Google Apps Script web app)
 * Serves Index.html and reads data from Google Sheets.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Atlas — Dynamic Data Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns {name, rows} where rows[0] is the header row.
 * url: a Google Sheet link (the tab in the link's #gid is used).
 *      Empty = the spreadsheet this script is bound to (first tab).
 */
function getSheetData(url) {
  var ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No sheet linked. Paste a Google Sheet link.');

  var sheet = ss.getSheets()[0];
  var gid = url && url.match(/[#&?]gid=(\d+)/);
  if (gid) {
    var match = ss.getSheets().filter(function (s) { return String(s.getSheetId()) === gid[1]; })[0];
    if (match) sheet = match;
  }

  var range = sheet.getDataRange();
  if (range.getNumRows() < 2) throw new Error('The sheet "' + sheet.getName() + '" has no data rows.');

  return {
    name: ss.getName() + ' / ' + sheet.getName(),
    rows: range.getDisplayValues()
  };
}
