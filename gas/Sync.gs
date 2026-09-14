function actionSync(params) {
  if (!CONFIG.SYNC_API_KEY) return errorResponse('GAS chưa cấu hình SYNC_API_KEY.', 'SYNC_NOT_CONFIGURED');
  if (!safeCompare(String(params.apiKey || ''), CONFIG.SYNC_API_KEY)) return errorResponse('Sai API key.', 'FORBIDDEN');

  var xuong = String(params.xuong || '').trim();
  var thang = String(params.thang || '').trim();
  var employees = Array.isArray(params.employees) ? params.employees : [];
  var payroll = Array.isArray(params.payroll) ? params.payroll : [];
  if (!xuong || !thang) return errorResponse('Thiếu thông tin xuong/thang.', 'MISSING_FIELDS');
  if (!/^\d{2}-\d{4}$/.test(thang)) return errorResponse('Tháng phải có dạng MM-YYYY.', 'INVALID_MONTH');
  if (!employees.length && !payroll.length) return errorResponse('Payload không có employees/payroll.', 'EMPTY_PAYLOAD');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var empResult = upsertEmployees(employees, xuong);
    var payrollResult = upsertPayroll(payroll, xuong, thang);
    updateSyncStatus(xuong, thang, payrollResult.count, 'OK', empResult.count + ' nhân viên, ' + payrollResult.count + ' phiếu lương');
    writeAudit('SYSTEM_VBA', 'SYNC', xuong + ' ' + thang, empResult.count + ' employees, ' + payrollResult.count + ' payroll rows');
    return successResponse({ message: 'Đồng bộ thành công.', employeesProcessed: empResult.count, payrollProcessed: payrollResult.count, thang: thang, xuong: xuong });
  } catch (err) {
    updateSyncStatus(xuong, thang, 0, 'ERROR', err.message || String(err));
    throw err;
  } finally { lock.releaseLock(); }
}

function upsertEmployees(employees, xuong) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var headers = requireHeaders(sheet, ['CCCD','MaNV','HoTen','Xuong','PhongBan','BoPhan','ChucVu','PasswordHash','PasswordSalt','MustChangePassword','Role','UpdatedAt']);
  if (sheet.getMaxRows() > 1) sheet.getRange(2, headers.CCCD, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  var rows = sheetToObjects(sheet), byMaNV = {};
  rows.forEach(function(r) { if (String(r.MaNV).trim()) byMaNV[String(r.MaNV).trim()] = r; });
  var now = formatDateVN(new Date()), count = 0;

  employees.forEach(function(e) {
    var maNV = String(e.maNV || '').trim();
    if (!maNV) return;
    var cccd = normalizeCCCD(e.cccd);
    var found = byMaNV[maNV];
    if (found) {
      sheet.getRange(found.__row, headers.CCCD).setNumberFormat('@').setValue(cccd);
      sheet.getRange(found.__row, headers.MaNV).setValue(maNV);
      sheet.getRange(found.__row, headers.HoTen).setValue(e.hoTen || found.HoTen || '');
      sheet.getRange(found.__row, headers.Xuong).setValue(xuong);
      sheet.getRange(found.__row, headers.PhongBan).setValue(e.phongBan || '');
      sheet.getRange(found.__row, headers.BoPhan).setValue(e.boPhan || '');
      sheet.getRange(found.__row, headers.ChucVu).setValue(e.chucVu || '');
      sheet.getRange(found.__row, headers.UpdatedAt).setValue(now);
    } else {
      var salt = generateSalt(), hash = hashPassword(maNV, salt);
      var row = [];
      sheet.getRange(1, headers.CCCD, Math.max(sheet.getMaxRows(), 2), 1).setNumberFormat('@');
      for (var i = 1; i <= sheet.getLastColumn(); i++) row.push('');
      row[headers.CCCD - 1] = cccd; row[headers.MaNV - 1] = maNV; row[headers.HoTen - 1] = e.hoTen || '';
      row[headers.Xuong - 1] = xuong; row[headers.PhongBan - 1] = e.phongBan || ''; row[headers.BoPhan - 1] = e.boPhan || '';
      row[headers.ChucVu - 1] = e.chucVu || ''; row[headers.PasswordHash - 1] = hash; row[headers.PasswordSalt - 1] = salt;
      row[headers.MustChangePassword - 1] = true; row[headers.Role - 1] = 'employee'; row[headers.UpdatedAt - 1] = now;
      sheet.appendRow(row);
    }
    count++;
  });
  return { count: count };
}

function upsertPayroll(payroll, xuong, thang) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL);
  var headers = requireHeaders(sheet, ['MaNV','HoTen','Xuong','Thang','UpdatedAt']);
  var rows = sheetToObjects(sheet), byKey = {};
  rows.forEach(function(r) { byKey[String(r.MaNV).trim() + '|' + String(r.Thang).trim()] = r; });
  var now = formatDateVN(new Date()), count = 0;
  payroll.forEach(function(p) {
    var maNV = String(p.maNV || '').trim(); if (!maNV) return;
    var key = maNV + '|' + thang;
    var data = buildPayrollRowData(p, xuong, thang, now, sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]);
    if (byKey[key]) sheet.getRange(byKey[key].__row, 1, 1, data.length).setValues([data]);
    else sheet.appendRow(data);
    count++;
  });
  return { count: count };
}

function buildPayrollRowData(p, xuong, thang, now, headers) {
  var map = {
    MaNV:p.maNV, HoTen:p.hoTen||'', Xuong:xuong, Thang:thang,
    LuongCoBan:p.luongCoBan||0, SoNgayLamViec:p.soNgayLamViec||0, SoNgayLe:p.soNgayLe||0,
    SoNgayNghiHuongLuong:p.soNgayNghiHuongLuong||0, SoNgayNghiKhongLuong:p.soNgayNghiKhongLuong||0,
    SoNgayNghiHuongLuongToiThieuVung:p.soNgayNghiHuongLuongToiThieuVung||0, LuongThang:p.luongThang||0,
    SoGioNgoaiGio:p.soGioNgoaiGio||0, SoGioNgayNghi:p.soGioNgayNghi||0, SoGioNgoaiGioNgayNghi:p.soGioNgoaiGioNgayNghi||0,
    SoGioTangCaDem:p.soGioTangCaDem||0, SoGioLamNgayLe:p.soGioLamNgayLe||0, SoGioNgoaiGioNgayLe:p.soGioNgoaiGioNgayLe||0,
    SoGioTangCaDemNgayLe:p.soGioTangCaDemNgayLe||0, SoNgayLamCaDem:p.soNgayLamCaDem||0, LuongNgoaiGio:p.luongNgoaiGio||0,
    TienKhac:p.tienKhac||0, TienKyLuat:p.tienKyLuat||0, TienGanBo2Nam:p.tienGanBo2Nam||0, TienGanBo5Nam:p.tienGanBo5Nam||0,
    TienGanBo10Nam:p.tienGanBo10Nam||0, TienNhaO:p.tienNhaO||0, TienDiLai:p.tienDiLai||0, TienThuongChuyenCan:p.tienThuongChuyenCan||0,
    HoaHongThuongVuotDinhMuc:p.hoaHongThuongVuotDinhMuc||0, TroCapThoiViecPhepNam:p.troCapThoiViecPhepNam||0,
    TongKhoanThuNhap:p.tongKhoanThuNhap||0, BHXH:p.bhxh||0, BHYT:p.bhyt||0, BHTN:p.bhtn||0, ThueThuNhap:p.thueThuNhap||0,
    TamUng:p.tamUng||0, KhauTruKhac:p.khauTruKhac||0, LuongThucLinh:p.luongThucLinh||0, UpdatedAt:now
  };
  return headers.map(function(h){ return Object.prototype.hasOwnProperty.call(map,h) ? map[h] : ''; });
}

function updateSyncStatus(xuong, thang, soDong, trangThai, ghiChu) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS, ['Xuong','Thang','LastSyncAt','SoDongDaXuLy','TrangThai','GhiChu']);
  var rows = sheetToObjects(sheet), found = rows.filter(function(r){return String(r.Xuong)===xuong && String(r.Thang)===thang;})[0];
  var now = formatDateVN(new Date()), data = [xuong,thang,now,soDong,trangThai,ghiChu||''];
  if (found) sheet.getRange(found.__row,1,1,6).setValues([data]); else sheet.appendRow(data);
}
