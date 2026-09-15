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


/**
 * Chạy một lần sau khi tạo project/deploy mới.
 * KHÔNG đặt secret trực tiếp trong source/GitHub.
 * Hãy vào Project Settings > Script properties và tạo đúng 3 tên:
 * SPREADSHEET_ID, SYNC_API_KEY, PASSWORD_PEPPER
 */
function validateConfiguration() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var missing = [];
  ['SPREADSHEET_ID','SYNC_API_KEY','PASSWORD_PEPPER'].forEach(function(k) {
    if (!props[k] || String(props[k]).trim() === '') missing.push(k);
  });
  if (missing.length) {
    throw new Error('Thiếu Script Properties: ' + missing.join(', ') + '. Vào Project Settings > Script properties và tạo đúng TÊN thuộc tính.');
  }
  try {
    var ss = SpreadsheetApp.openById(String(props.SPREADSHEET_ID).trim());
    return {ok:true, spreadsheetIdConfigured:true, spreadsheetName:ss.getName(),
      syncApiKeyConfigured:true, passwordPepperConfigured:true};
  } catch (e) {
    throw new Error('SPREADSHEET_ID đã có nhưng không mở được Google Sheet. Kiểm tra ID và quyền truy cập của tài khoản chạy Apps Script. Chi tiết: ' + e.message);
  }
}

function setConfigurationFromPrompt() {
  var ui = SpreadsheetApp.getUi();
  var ssId = ui.prompt('SPREADSHEET_ID', 'Nhập ID Google Sheet:', ui.ButtonSet.OK_CANCEL);
  if (ssId.getSelectedButton() !== ui.Button.OK) return 'Đã hủy.';
  var api = ui.prompt('SYNC_API_KEY', 'Nhập API key dùng cho VBA:', ui.ButtonSet.OK_CANCEL);
  if (api.getSelectedButton() !== ui.Button.OK) return 'Đã hủy.';
  var pepper = ui.prompt('PASSWORD_PEPPER', 'Nhập PASSWORD_PEPPER:', ui.ButtonSet.OK_CANCEL);
  if (pepper.getSelectedButton() !== ui.Button.OK) return 'Đã hủy.';
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    SPREADSHEET_ID: String(ssId.getResponseText()).trim(),
    SYNC_API_KEY: String(api.getResponseText()).trim(),
    PASSWORD_PEPPER: String(pepper.getResponseText()).trim()
  }, true);
  return validateConfiguration();
}
