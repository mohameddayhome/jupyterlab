/**
 * Atlas — Dynamic Dashboard (Google Apps Script web app)
 * Serves Index.html and reads data from Google Sheets in chunks.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('أطلس — لوحة البيانات التفاعلية')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * url: a Google Sheet link (the tab in the link's #gid is used).
 *      Empty = the spreadsheet this script is bound to (first tab).
 */
function resolveSheet_(url, gid) {
  var ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('لا يوجد جدول مرتبط. الصق رابط Google Sheet.');

  if (gid === undefined || gid === null) {
    var m = url && url.match(/[#&?]gid=(\d+)/);
    gid = m ? m[1] : null;
  }
  var sheet = ss.getSheets()[0];
  if (gid !== null) {
    var match = ss.getSheets().filter(function (s) { return String(s.getSheetId()) === String(gid); })[0];
    if (match) sheet = match;
  }
  return { ss: ss, sheet: sheet };
}

/** Size of the sheet, so the page can fetch it in chunks. */
function getSheetInfo(url) {
  var r = resolveSheet_(url);
  return {
    name: r.ss.getName() + ' / ' + r.sheet.getName(),
    gid: r.sheet.getSheetId(),
    rows: r.sheet.getLastRow(),
    cols: r.sheet.getLastColumn()
  };
}

/** Rows [start, start + count) as a JSON string (row 1 is the header). */
function getSheetChunk(url, gid, start, count) {
  var sheet = resolveSheet_(url, gid).sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var n = Math.min(count, lastRow - start + 1);
  if (n <= 0 || lastCol <= 0) return '[]';
  return JSON.stringify(sheet.getRange(start, 1, n, lastCol).getDisplayValues());
}
