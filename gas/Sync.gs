/**
 * Sync.gs
 * -----------------------------------------------------------------------
 * Endpoint nhận dữ liệu do VBA đẩy lên (từ file Excel .xlsb). Bảo vệ bằng
 * SYNC_API_KEY (không dùng session token nhân viên/admin).
 *
 * Payload dự kiến từ VBA (xem vba/SyncPayroll.bas):
 * {
 *   apiKey: "...",
 *   xuong: "Snack" | "Flexible",
 *   thang: "08-2026",
 *   employees: [ { cccd, maNV, hoTen, phongBan, boPhan, chucVu }, ... ],
 *   payroll:   [ { maNV, hoTen, ...tất cả cột lương... }, ... ]
 * }
 * -----------------------------------------------------------------------
 */

function actionSync(params) {
  if (!params.apiKey || !safeCompare(String(params.apiKey), CONFIG.SYNC_API_KEY)) {
    return errorResponse('Sai API key.', 'FORBIDDEN');
  }

  var xuong = String(params.xuong || '').trim();
  var thang = String(params.thang || '').trim();
  if (!xuong || !thang) {
    return errorResponse('Thiếu thông tin xuong/thang.', 'MISSING_FIELDS');
  }

  var employees = params.employees || [];
  var payroll = params.payroll || [];

  var empResult = upsertEmployees(employees, xuong);
  var payrollResult = upsertPayroll(payroll, xuong, thang);

  updateSyncStatus(xuong, thang, empResult.count + payrollResult.count, 'OK',
    empResult.count + ' nhân viên, ' + payrollResult.count + ' phiếu lương');

  writeAudit('SYSTEM_VBA', 'SYNC', xuong + ' ' + thang,
    empResult.count + ' employees, ' + payrollResult.count + ' payroll rows');

  return successResponse({
    message: 'Đồng bộ thành công.',
    employeesProcessed: empResult.count,
    payrollProcessed: payrollResult.count
  });
}

/**
 * Thêm mới hoặc cập nhật nhân viên (dò theo MaNV). Nhân viên mới sẽ được đặt
 * mật khẩu ban đầu = Mã nhân viên, bắt buộc đổi mật khẩu lần đầu.
 */
function upsertEmployees(employees, xuong) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var existing = sheetToObjects(sheet);
  var byMaNV = {};
  existing.forEach(function (r) { byMaNV[String(r.MaNV)] = r; });

  var now = formatDateVN(new Date());
  var count = 0;

  employees.forEach(function (e) {
    if (!e.maNV) return;
    var maNV = String(e.maNV).trim();
    var cccd = normalizeCCCD(e.cccd);
    var found = byMaNV[maNV];

    if (found) {
      // Cập nhật thông tin, KHÔNG đụng vào PasswordHash/Salt/MustChangePassword
      setRowField(sheet, headers, found.__row, 'CCCD', "'" + cccd); // giữ số 0 đầu
      setRowField(sheet, headers, found.__row, 'HoTen', e.hoTen || found.HoTen);
      setRowField(sheet, headers, found.__row, 'Xuong', xuong);
      setRowField(sheet, headers, found.__row, 'PhongBan', e.phongBan || '');
      setRowField(sheet, headers, found.__row, 'BoPhan', e.boPhan || '');
      setRowField(sheet, headers, found.__row, 'ChucVu', e.chucVu || '');
      setRowField(sheet, headers, found.__row, 'UpdatedAt', now);
    } else {
      var salt = generateSalt();
      var hash = hashPassword(maNV, salt); // mật khẩu ban đầu = mã nhân viên
      var newRow = [];
      headers.forEach(function (h) {
        switch (h) {
          case 'CCCD': newRow.push("'" + cccd); break;
          case 'MaNV': newRow.push(maNV); break;
          case 'HoTen': newRow.push(e.hoTen || ''); break;
          case 'Xuong': newRow.push(xuong); break;
          case 'PhongBan': newRow.push(e.phongBan || ''); break;
          case 'BoPhan': newRow.push(e.boPhan || ''); break;
          case 'ChucVu': newRow.push(e.chucVu || ''); break;
          case 'PasswordHash': newRow.push(hash); break;
          case 'PasswordSalt': newRow.push(salt); break;
          case 'MustChangePassword': newRow.push(true); break;
          case 'Role': newRow.push('employee'); break;
          case 'UpdatedAt': newRow.push(now); break;
          default: newRow.push('');
        }
      });
      sheet.appendRow(newRow);
    }
    count++;
  });

  return { count: count };
}

/**
 * Thêm mới hoặc cập nhật (upsert theo MaNV+Thang) 1 dòng phiếu lương.
 */
function upsertPayroll(payroll, xuong, thang) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var existing = sheetToObjects(sheet);
  var byKey = {};
  existing.forEach(function (r) {
    byKey[String(r.MaNV) + '|' + r.Thang] = r;
  });

  var now = formatDateVN(new Date());
  var count = 0;

  payroll.forEach(function (p) {
    if (!p.maNV) return;
    var key = String(p.maNV).trim() + '|' + thang;
    var found = byKey[key];
    var rowData = buildPayrollRowData(p, xuong, thang, now, headers);

    if (found) {
      sheet.getRange(found.__row, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    count++;
  });

  return { count: count };
}

function buildPayrollRowData(p, xuong, thang, now, headers) {
  var map = {
    MaNV: p.maNV, HoTen: p.hoTen || '', Xuong: xuong, Thang: thang,
    LuongCoBan: p.luongCoBan || 0,
    SoNgayLamViec: p.soNgayLamViec || 0,
    SoNgayLe: p.soNgayLe || 0,
    SoNgayNghiHuongLuong: p.soNgayNghiHuongLuong || 0,
    SoNgayNghiKhongLuong: p.soNgayNghiKhongLuong || 0,
    SoNgayNghiHuongLuongToiThieuVung: p.soNgayNghiHuongLuongToiThieuVung || 0,
    LuongThang: p.luongThang || 0,
    SoGioNgoaiGio: p.soGioNgoaiGio || 0,
    SoGioNgayNghi: p.soGioNgayNghi || 0,
    SoGioNgoaiGioNgayNghi: p.soGioNgoaiGioNgayNghi || 0,
    SoGioTangCaDem: p.soGioTangCaDem || 0,
    SoGioLamNgayLe: p.soGioLamNgayLe || 0,
    SoGioNgoaiGioNgayLe: p.soGioNgoaiGioNgayLe || 0,
    SoGioTangCaDemNgayLe: p.soGioTangCaDemNgayLe || 0,
    SoNgayLamCaDem: p.soNgayLamCaDem || 0,
    LuongNgoaiGio: p.luongNgoaiGio || 0,
    TienKhac: p.tienKhac || 0,
    TienKyLuat: p.tienKyLuat || 0,
    TienGanBo2Nam: p.tienGanBo2Nam || 0,
    TienGanBo5Nam: p.tienGanBo5Nam || 0,
    TienGanBo10Nam: p.tienGanBo10Nam || 0,
    TienNhaO: p.tienNhaO || 0,
    TienDiLai: p.tienDiLai || 0,
    TienThuongChuyenCan: p.tienThuongChuyenCan || 0,
    HoaHongThuongVuotDinhMuc: p.hoaHongThuongVuotDinhMuc || 0,
    TroCapThoiViecPhepNam: p.troCapThoiViecPhepNam || 0,
    TongKhoanThuNhap: p.tongKhoanThuNhap || 0,
    BHXH: p.bhxh || 0,
    BHYT: p.bhyt || 0,
    BHTN: p.bhtn || 0,
    ThueThuNhap: p.thueThuNhap || 0,
    TamUng: p.tamUng || 0,
    KhauTruKhac: p.khauTruKhac || 0,
    LuongThucLinh: p.luongThucLinh || 0,
    UpdatedAt: now
  };
  return headers.map(function (h) { return map.hasOwnProperty(h) ? map[h] : ''; });
}

function setRowField(sheet, headers, row, fieldName, value) {
  var col = headers.indexOf(fieldName) + 1;
  if (col > 0) sheet.getRange(row, col).setValue(value);
}

function updateSyncStatus(xuong, thang, soDong, trangThai, ghiChu) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS);
  var rows = sheetToObjects(sheet);
  var found = rows.filter(function (r) { return r.Xuong === xuong && r.Thang === thang; })[0];
  var now = formatDateVN(new Date());
  if (found) {
    sheet.getRange(found.__row, 1, 1, 6).setValues([[xuong, thang, now, soDong, trangThai, ghiChu]]);
  } else {
    sheet.appendRow([xuong, thang, now, soDong, trangThai, ghiChu]);
  }
}
