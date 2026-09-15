/**
 * Utils.gs
 * -----------------------------------------------------------------------
 * Các hàm tiện ích dùng chung: format số, format tiền, JSON response, ghi log...
 * -----------------------------------------------------------------------
 */

/**
 * Trả về JSON response chuẩn cho web app.
 * ok: true/false, data: payload, error: thông báo lỗi (nếu có)
 */
function jsonOutput(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function successResponse(data) {
  return jsonOutput({ ok: true, data: data });
}

function errorResponse(message, code) {
  return jsonOutput({ ok: false, error: message, code: code || 'ERROR' });
}

/**
 * Format số tiền VNĐ, ví dụ 1234567 -> "1.234.567 đ"
 */
function formatCurrencyVND(value) {
  var n = Number(value) || 0;
  var rounded = Math.round(n);
  var parts = Math.abs(rounded).toString().split('').reverse();
  var withDots = [];
  for (var i = 0; i < parts.length; i++) {
    if (i > 0 && i % 3 === 0) withDots.push('.');
    withDots.push(parts[i]);
  }
  var formatted = withDots.reverse().join('');
  return (rounded < 0 ? '-' : '') + formatted + ' đ';
}

/**
 * Format số ngày/giờ: lấy 2 số thập phân nếu là số lẻ, giữ nguyên nếu là số nguyên.
 * Ví dụ: 26 -> "26", 25.5 -> "25.50", 12.333 -> "12.33"
 */
function formatDaysHours(value) {
  var n = Number(value);
  if (isNaN(n)) return '';
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(2);
}

/**
 * Format ngày giờ theo dd/MM/yyyy HH:mm:ss (theo timezone của Spreadsheet).
 */
function formatDateVN(date) {
  var tz = getSS().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'dd/MM/yyyy HH:mm:ss');
}

/**
 * Ghi 1 dòng vào sheet AuditLog.
 */
function writeAudit(actor, action, target, detail) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_AUDIT, ['Timestamp', 'Actor', 'Action', 'Target', 'Detail']);
  sheet.appendRow([formatDateVN(new Date()), actor, action, target, detail || '']);
}

/**
 * So sánh 2 chuỗi theo kiểu "constant time" để hạn chế timing attack khi so
 * sánh hash mật khẩu / token. Apps Script không có hàm built-in cho việc này
 * nên tự triển khai đơn giản.
 */
function safeCompare(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  if (a.length !== b.length) return false;
  var result = 0;
  for (var i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

/**
 * Chuẩn hoá số Căn cước: loại bỏ khoảng trắng, dấu nháy đầu (nếu VBA gửi kèm),
 * chỉ giữ lại chữ số.
 */
function normalizeCCCD(value) {
  return String(value || '').replace(/[^0-9]/g, '').trim();
}

/**
 * Đọc toàn bộ dữ liệu của 1 sheet thành mảng object (dùng header dòng 1).
 */
function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    // Bỏ qua dòng trống hoàn toàn
    var isEmpty = row.every(function (c) { return c === '' || c === null; });
    if (isEmpty) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    obj.__row = i + 1; // số dòng thực tế trong sheet (1-based), phục vụ update
    rows.push(obj);
  }
  return rows;
}

/**
 * Thêm CORS header cho phép GitHub Pages gọi tới (Apps Script Web App vốn đã
 * cho phép CORS ở mức response body/text, nhưng nếu bạn dùng doOptions thì có
 * thể mở rộng thêm ở đây nếu cần).
 */
function withCors(output) {
  // ContentService không cho set custom header trực tiếp trong bản Apps Script
  // hiện tại; CORS cho JSON POST/GET đơn giản (text/plain hoặc application/json
  // không có custom header) thường không bị chặn bởi trình duyệt khi gọi từ
  // GitHub Pages tới *.googleusercontent.com. Nếu gặp lỗi CORS, xem docs/GAS_DEPLOY.md.
  return output;
}
