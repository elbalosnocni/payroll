/**
 * Sync.gs - VBA -> Google Sheets
 * IMPORTANT: all writes are batched. Do NOT use appendRow/setValue in a loop.
 */
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
  // The VBA requests are sequential. 2 minutes is enough after batching, and prevents a stale request from blocking forever.
  try {
    lock.waitLock(120000);
  } catch (e) {
    return errorResponse('Không lấy được khóa đồng bộ sau 120 giây. Có thể một lần đồng bộ trước đang chạy. Hãy chờ rồi chạy lại.', 'LOCK_TIMEOUT');
  }

  try {
    var empResult = upsertEmployeesBatch(employees, xuong);
    var payrollResult = upsertPayrollBatch(payroll, xuong, thang);
    updateSyncStatus(xuong, thang, payrollResult.count, 'OK', empResult.count + ' nhân viên, ' + payrollResult.count + ' phiếu lương');
    writeAudit('SYSTEM_VBA', 'SYNC', xuong + ' ' + thang, empResult.count + ' employees, ' + payrollResult.count + ' payroll rows');
    SpreadsheetApp.flush();
    return successResponse({
      message: 'Đồng bộ thành công.',
      employeesProcessed: empResult.count,
      payrollProcessed: payrollResult.count,
      thang: thang,
      xuong: xuong,
      durationMs: new Date().getTime() - Number(params.clientStartedAt || new Date().getTime())
    });
  } catch (err) {
    try { updateSyncStatus(xuong, thang, 0, 'ERROR', err.message || String(err)); } catch (_) {}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function upsertEmployeesBatch(employees, xuong) {
  var headersList = ['CCCD','MaNV','HoTen','Xuong','PhongBan','BoPhan','ChucVu','PasswordHash','PasswordSalt','MustChangePassword','Role','UpdatedAt'];
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES, headersList);
  var headers = requireHeaders(sheet, headersList);
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), headersList.length);
  var values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var byMa = {};
  for (var i = 0; i < values.length; i++) {
    var ma = String(values[i][headers.MaNV - 1] == null ? '' : values[i][headers.MaNV - 1]).trim();
    if (ma && !byMa[ma]) byMa[ma] = i;
  }

  var now = formatDateVN(new Date());
  var processed = 0;
  employees.forEach(function(e) {
    var maNV = String(e.maNV || '').trim();
    if (!maNV) return;
    var idx = Object.prototype.hasOwnProperty.call(byMa, maNV) ? byMa[maNV] : -1;
    var row;
    if (idx >= 0) {
      row = values[idx];
      // Existing password/salt/must-change/role are intentionally preserved.
      row[headers.CCCD - 1] = normalizeCCCD(e.cccd);
      row[headers.MaNV - 1] = maNV;
      row[headers.HoTen - 1] = e.hoTen || row[headers.HoTen - 1] || '';
      row[headers.Xuong - 1] = xuong;
      row[headers.PhongBan - 1] = e.phongBan || '';
      row[headers.BoPhan - 1] = e.boPhan || '';
      row[headers.ChucVu - 1] = e.chucVu || '';
      row[headers.UpdatedAt - 1] = now;
    } else {
      row = new Array(lastCol).fill('');
      row[headers.CCCD - 1] = normalizeCCCD(e.cccd);
      row[headers.MaNV - 1] = maNV;
      row[headers.HoTen - 1] = e.hoTen || '';
      row[headers.Xuong - 1] = xuong;
      row[headers.PhongBan - 1] = e.phongBan || '';
      row[headers.BoPhan - 1] = e.boPhan || '';
      row[headers.ChucVu - 1] = e.chucVu || '';
      var salt = generateSalt();
      row[headers.PasswordSalt - 1] = salt;
      row[headers.PasswordHash - 1] = hashPassword(maNV, salt);
      row[headers.MustChangePassword - 1] = true;
      row[headers.Role - 1] = 'employee';
      row[headers.UpdatedAt - 1] = now;
      byMa[maNV] = values.length;
      values.push(row);
    }
    processed++;
  });

  if (values.length) {
    sheet.getRange(2, 1, values.length, lastCol).setValues(values);
    sheet.getRange(2, headers.CCCD, values.length, 1).setNumberFormat('@');
    sheet.getRange(2, headers.MaNV, values.length, 1).setNumberFormat('@');
  }
  return {count: processed};
}

function upsertPayrollBatch(payroll, xuong, thang) {
  var headersList = ['MaNV','HoTen','Xuong','Thang','LuongCoBan','SoNgayLamViec','SoNgayLe','SoNgayNghiHuongLuong','SoNgayNghiKhongLuong','SoNgayNghiHuongLuongToiThieuVung','LuongThang','SoGioNgoaiGio','SoGioNgayNghi','SoGioNgoaiGioNgayNghi','SoGioTangCaDem','SoGioLamNgayLe','SoGioNgoaiGioNgayLe','SoGioTangCaDemNgayLe','SoNgayLamCaDem','LuongNgoaiGio','TienKhac','TienKyLuat','TienGanBo2Nam','TienGanBo5Nam','TienGanBo10Nam','TienNhaO','TienDiLai','TienThuongChuyenCan','HoaHongThuongVuotDinhMuc','TroCapThoiViecPhepNam','TongKhoanThuNhap','BHXH','BHYT','BHTN','ThueThuNhap','TamUng','KhauTruKhac','LuongThucLinh','UpdatedAt'];
  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL, headersList);
  var headers = requireHeaders(sheet, headersList);
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), headersList.length);
  var values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var byKey = {};
  for (var i = 0; i < values.length; i++) {
    var ma = String(values[i][headers.MaNV - 1] == null ? '' : values[i][headers.MaNV - 1]).trim();
    var th = String(values[i][headers.Thang - 1] == null ? '' : values[i][headers.Thang - 1]).trim();
    if (ma && th && !byKey[ma + '|' + th]) byKey[ma + '|' + th] = i;
  }
  var now = formatDateVN(new Date());
  var processed = 0;
  payroll.forEach(function(p) {
    var maNV = String(p.maNV || '').trim();
    if (!maNV) return;
    var key = maNV + '|' + thang;
    var data = buildPayrollRowData(p, xuong, thang, now, headersList);
    var idx = Object.prototype.hasOwnProperty.call(byKey, key) ? byKey[key] : -1;
    if (idx >= 0) values[idx] = data;
    else { byKey[key] = values.length; values.push(data); }
    processed++;
  });
  if (values.length) sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  return {count: processed};
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
  var lastRow = sheet.getLastRow();
  var values = lastRow >= 2 ? sheet.getRange(2,1,lastRow-1,6).getValues() : [];
  var found = -1;
  for (var i=0;i<values.length;i++) {
    if (String(values[i][0]).trim()===xuong && String(values[i][1]).trim()===thang) { found=i; break; }
  }
  var data = [xuong,thang,formatDateVN(new Date()),soDong,trangThai,ghiChu||''];
  if (found >= 0) sheet.getRange(found+2,1,1,6).setValues([data]);
  else sheet.getRange(Math.max(2,lastRow+1),1,1,6).setValues([data]);
}
