var CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  SYNC_API_KEY: PropertiesService.getScriptProperties().getProperty('SYNC_API_KEY') || '',
  PASSWORD_PEPPER: PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '',
  SHEET_EMPLOYEES: 'Employees',
  SHEET_PAYROLL: 'Payroll',
  SHEET_AUDIT: 'AuditLog',
  SHEET_SYNC_STATUS: 'SyncStatus',
  SESSION_TTL_SEC: 8 * 60 * 60,
  HASH_ITERATIONS: 12000,
  MAX_LOGIN_ATTEMPTS: 8,
  LOGIN_WINDOW_SEC: 15 * 60
};

function getSS() {
  if (!CONFIG.SPREADSHEET_ID) throw new Error('Chưa cấu hình SPREADSHEET_ID.');
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(name, headers) {
  var ss = getSS();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sh;
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

  var emp = getSS().getSheetByName(CONFIG.SHEET_EMPLOYEES);
  emp.getRange('A:A').setNumberFormat('@');
  emp.getRange('L:L').setNumberFormat('@');
  emp.setFrozenRows(1);
  Logger.log('Payroll system setup OK.');
}

function setInitialAdmin(maNV) {
  maNV = String(maNV || '').trim();
  if (!maNV) throw new Error('Thiếu Mã NV admin.');
  var sh = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var rows = sheetToObjects(sh);
  var found = rows.filter(function(r){ return String(r.MaNV) === maNV; })[0];
  if (!found) throw new Error('Không tìm thấy MaNV.');
  var salt = generateSalt();
  setField(sh, found.__row, 'PasswordHash', hashPassword(maNV, salt));
  setField(sh, found.__row, 'PasswordSalt', salt);
  setField(sh, found.__row, 'MustChangePassword', true);
  setField(sh, found.__row, 'Role', 'ADMIN');
  setField(sh, found.__row, 'UpdatedAt', formatDateVN(new Date()));
}
