function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function successResponse(data) { return jsonOutput({ ok: true, data: data }); }
function errorResponse(message, code) { return jsonOutput({ ok: false, error: message, code: code || 'ERROR' }); }

function formatCurrencyVND(value) {
  var n = Number(value);
  if (!isFinite(n)) n = 0;
  var rounded = Math.round(n);
  var s = Math.abs(rounded).toString();
  var out = '';
  for (var i = 0; i < s.length; i++) {
    if (i > 0 && (s.length - i) % 3 === 0) out += '.';
    out += s.charAt(i);
  }
  return (rounded < 0 ? '-' : '') + out + ' đ';
}

function formatDaysHours(value) {
  var n = Number(value);
  if (!isFinite(n)) return '';
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}

function formatDateVN(date) {
  return Utilities.formatDate(date, getSS().getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function writeAudit(actor, action, target, detail) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_AUDIT, ['Timestamp','Actor','Action','Target','Detail']);
  sheet.appendRow([formatDateVN(new Date()), actor || '', action || '', target || '', detail || '']);
}

function safeCompare(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  var r = 0;
  for (var i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function normalizeCCCD(value) {
  return String(value == null ? '' : value).replace(/[^0-9]/g, '');
}

function isTrue(value) {
  return value === true || String(value).toLowerCase() === 'true' || String(value) === '1';
}

function sheetToObjects(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var empty = row.every(function(c) { return c === '' || c === null; });
    if (empty) continue;
    var obj = {};
    headers.forEach(function(h, j) { if (h) obj[h] = row[j]; });
    obj.__row = i + 1;
    rows.push(obj);
  }
  return rows;
}

function headerMap(sheet) {
  var n = sheet.getLastColumn();
  if (!n) return {};
  var headers = sheet.getRange(1, 1, 1, n).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) { map[String(h).trim()] = i + 1; });
  return map;
}

function requireHeaders(sheet, required) {
  var map = headerMap(sheet);
  var missing = required.filter(function(h) { return !map[h]; });
  if (missing.length) throw new Error('Sheet ' + sheet.getName() + ' thiếu cột: ' + missing.join(', '));
  return map;
}
