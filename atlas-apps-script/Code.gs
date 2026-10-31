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

/**
 * Saved lists (calculated columns, layout templates) as JSON strings in script properties.
 * Values are split into chunks because one property holds at most ~9 KB.
 */
var STORE_KEYS_ = { 'calc-columns': 'atlasCalcColumns', 'layouts': 'atlasLayouts' };
var CHUNK_ = 8000;

function storeKey_(name) {
  var key = STORE_KEYS_[name];
  if (!key) throw new Error('Unknown store: ' + name);
  return key;
}

function getStored(name) {
  var key = storeKey_(name);
  var props = PropertiesService.getScriptProperties();
  var count = Number(props.getProperty(key + '_n') || 0);
  if (!count) return props.getProperty(key) || '[]'; // single value (older saves)
  var parts = [];
  for (var i = 0; i < count; i++) parts.push(props.getProperty(key + '_' + i) || '');
  return parts.join('');
}

function saveStored(name, json) {
  var key = storeKey_(name);
  JSON.parse(json); // reject malformed input
  var props = PropertiesService.getScriptProperties();
  var oldCount = Number(props.getProperty(key + '_n') || 0);
  var values = {};
  var count = Math.ceil(json.length / CHUNK_) || 1;
  for (var i = 0; i < count; i++) values[key + '_' + i] = json.slice(i * CHUNK_, (i + 1) * CHUNK_);
  values[key + '_n'] = String(count);
  props.setProperties(values);
  for (var j = count; j < oldCount; j++) props.deleteProperty(key + '_' + j);
  props.deleteProperty(key);
}
