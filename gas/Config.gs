/**
 * Config.gs
 * -----------------------------------------------------------------------
 * Tất cả cấu hình tập trung tại đây. KHÔNG hard-code các giá trị nhạy cảm
 * (API key, Spreadsheet ID...) trực tiếp trong code khi share cho người khác;
 * nên lưu trong Script Properties (File > Project properties > Script properties)
 * và đọc ra bằng PropertiesService như bên dưới.
 * -----------------------------------------------------------------------
 */

var CONFIG = {
  // ID của Google Sheet dùng làm database (lấy trong URL của sheet)
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || 'PUT_YOUR_SPREADSHEET_ID_HERE',

  // Khóa bí mật dùng để VBA xác thực khi đẩy dữ liệu lên (endpoint /sync)
  // Đặt trong Script Properties: SYNC_API_KEY
  SYNC_API_KEY: PropertiesService.getScriptProperties().getProperty('SYNC_API_KEY') || 'CHANGE_ME_SYNC_KEY',

  // Salt phụ (pepper) cộng thêm khi hash mật khẩu, tăng độ an toàn.
  // Đặt trong Script Properties: PASSWORD_PEPPER
  PASSWORD_PEPPER: PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || 'CHANGE_ME_PEPPER',

  // Tên các sheet trong Spreadsheet
  SHEET_EMPLOYEES: 'Employees',
  SHEET_PAYROLL: 'Payroll',
  SHEET_AUDIT: 'AuditLog',
  SHEET_SYNC_STATUS: 'SyncStatus',

  // Thời hạn 1 session (giây) - 8 tiếng
  SESSION_DURATION_SEC: 8 * 60 * 60,

  // Số vòng lặp băm mật khẩu (PBKDF2-like, tăng độ khó brute-force)
  HASH_ITERATIONS: 10000,

  // Danh sách domain được phép gọi API (CORS). '*' = cho phép tất cả.
  // Nên đổi thành domain GitHub Pages thực tế của bạn, ví dụ:
  // 'https://tenban.github.io'
  ALLOWED_ORIGIN: '*'
};

/**
 * Lấy đối tượng Spreadsheet dùng chung trong toàn bộ project.
 */
function getSS() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Lấy 1 sheet theo tên, tạo mới với header nếu chưa tồn tại.
 */
function getOrCreateSheet(name, headers) {
  var ss = getSS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Khởi tạo cấu trúc sheet cần thiết (chạy 1 lần thủ công từ Apps Script Editor:
 * chọn hàm setupSheets rồi bấm Run).
 */
function setupSheets() {
  getOrCreateSheet(CONFIG.SHEET_EMPLOYEES, [
    'CCCD', 'MaNV', 'HoTen', 'Xuong', 'PhongBan', 'BoPhan', 'ChucVu',
    'PasswordHash', 'PasswordSalt', 'MustChangePassword', 'Role',
    'UpdatedAt'
  ]);
  getOrCreateSheet(CONFIG.SHEET_PAYROLL, [
    'MaNV', 'HoTen', 'Xuong', 'Thang',
    'LuongCoBan', 'SoNgayLamViec', 'SoNgayLe', 'SoNgayNghiHuongLuong',
    'SoNgayNghiKhongLuong', 'SoNgayNghiHuongLuongToiThieuVung', 'LuongThang',
    'SoGioNgoaiGio', 'SoGioNgayNghi', 'SoGioNgoaiGioNgayNghi', 'SoGioTangCaDem',
    'SoGioLamNgayLe', 'SoGioNgoaiGioNgayLe', 'SoGioTangCaDemNgayLe',
    'SoNgayLamCaDem', 'LuongNgoaiGio',
    'TienKhac', 'TienKyLuat', 'TienGanBo2Nam', 'TienGanBo5Nam', 'TienGanBo10Nam',
    'TienNhaO', 'TienDiLai', 'TienThuongChuyenCan', 'HoaHongThuongVuotDinhMuc',
    'TroCapThoiViecPhepNam', 'TongKhoanThuNhap',
    'BHXH', 'BHYT', 'BHTN', 'ThueThuNhap', 'TamUng', 'KhauTruKhac',
    'LuongThucLinh', 'UpdatedAt'
  ]);
  getOrCreateSheet(CONFIG.SHEET_AUDIT, [
    'Timestamp', 'Actor', 'Action', 'Target', 'Detail'
  ]);
  getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS, [
    'Xuong', 'Thang', 'LastSyncAt', 'SoDongDaXuLy', 'TrangThai', 'GhiChu'
  ]);
  Logger.log('Setup hoàn tất.');
}
