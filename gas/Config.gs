/**
 * Config.gs - cấu hình tập trung.
 *
 * BẮT BUỘC cấu hình 3 Script Properties trong Apps Script:
 *   SPREADSHEET_ID
 *   SYNC_API_KEY
 *   PASSWORD_PEPPER
 *
 * Không đặt dữ liệu thật vào GitHub.
 */
var CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  SYNC_API_KEY: PropertiesService.getScriptProperties().getProperty('SYNC_API_KEY') || '',
  PASSWORD_PEPPER: PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '',
  SHEET_EMPLOYEES: 'Employees',
  SHEET_PAYROLL: 'Payroll',
  SHEET_AUDIT: 'AuditLog',
  SHEET_SYNC_STATUS: 'SyncStatus',
  SESSION_DURATION_SEC: 8 * 60 * 60,
  HASH_ITERATIONS: 10000
};

function getSS() {
  if (!CONFIG.SPREADSHEET_ID) throw new Error('Chưa cấu hình Script Property SPREADSHEET_ID.');
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(name, headers) {
  var ss = getSS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (headers && headers.length) {
    var lastCol = sheet.getLastColumn();
    var current = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var needsHeader = current.length === 0 || current.every(function(v) { return String(v).trim() === ''; });
    if (needsHeader) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function setupSheets() {
  getOrCreateSheet(CONFIG.SHEET_EMPLOYEES, [
    'CCCD','MaNV','HoTen','Xuong','PhongBan','BoPhan','ChucVu',
    'PasswordHash','PasswordSalt','MustChangePassword','Role','UpdatedAt'
  ]);
  getOrCreateSheet(CONFIG.SHEET_PAYROLL, [
    'MaNV','HoTen','Xuong','Thang','LuongCoBan','SoNgayLamViec','SoNgayLe',
    'SoNgayNghiHuongLuong','SoNgayNghiKhongLuong','SoNgayNghiHuongLuongToiThieuVung',
    'LuongThang','SoGioNgoaiGio','SoGioNgayNghi','SoGioNgoaiGioNgayNghi',
    'SoGioTangCaDem','SoGioLamNgayLe','SoGioNgoaiGioNgayLe','SoGioTangCaDemNgayLe',
    'SoNgayLamCaDem','LuongNgoaiGio','TienKhac','TienKyLuat','TienGanBo2Nam',
    'TienGanBo5Nam','TienGanBo10Nam','TienNhaO','TienDiLai','TienThuongChuyenCan',
    'HoaHongThuongVuotDinhMuc','TroCapThoiViecPhepNam','TongKhoanThuNhap',
    'BHXH','BHYT','BHTN','ThueThuNhap','TamUng','KhauTruKhac','LuongThucLinh','UpdatedAt'
  ]);
  getOrCreateSheet(CONFIG.SHEET_AUDIT, ['Timestamp','Actor','Action','Target','Detail']);
  getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS, ['Xuong','Thang','LastSyncAt','SoDongDaXuLy','TrangThai','GhiChu']);
  return 'OK';
}

function checkConfiguration() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return {
    spreadsheetIdConfigured: !!props.SPREADSHEET_ID,
    syncApiKeyConfigured: !!props.SYNC_API_KEY,
    passwordPepperConfigured: !!props.PASSWORD_PEPPER,
    spreadsheetIdMasked: props.SPREADSHEET_ID ? String(props.SPREADSHEET_ID).slice(0, 6) + '...' : '',
    spreadsheetName: props.SPREADSHEET_ID ? getSS().getName() : ''
  };
}
