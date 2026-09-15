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
 *
 * TỐI ƯU HIỆU NĂNG (quan trọng):
 * - Bản trước ghi từng ô/từng dòng bằng setValue()/appendRow() lặp trong vòng
 *   for => với ~50-60 nhân viên có thể tốn tới hàng trăm lệnh gọi Range API,
 *   khiến 1 lần đồng bộ chạy 1-2 phút. Hệ quả: (1) request HTTP phía VBA bị
 *   timeout dù server vẫn đang chạy ngầm, và (2) khi xưởng thứ 2 gọi lên gần
 *   như cùng lúc, request đó phải CHỜ (hoặc bị từ chối SYNC_BUSY) vì request
 *   trước vẫn đang giữ khóa/ghi dữ liệu.
 * - Bản này đọc/ghi theo LÔ: mỗi lần đồng bộ chỉ cần tối đa 2 lệnh getValues()
 *   và 2 lệnh setValues() (không phụ thuộc số lượng nhân viên), nên toàn bộ
 *   quá trình chỉ mất vài giây => loại bỏ gốc rễ của cả 2 lỗi timeout/BUSY.
 * - Đồng thời bọc toàn bộ thao tác ghi trong LockService để đảm bảo 2 xưởng
 *   gửi lên cùng lúc không bao giờ ghi đè lẫn nhau; nếu thực sự phải chờ quá
 *   SYNC_LOCK_WAIT_MS thì trả lỗi SYNC_BUSY rõ ràng để VBA tự động thử lại.
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

  // Khóa toàn cục: đảm bảo tại 1 thời điểm chỉ có 1 request /sync được ghi
  // vào Sheet, tránh 2 xưởng (hoặc 2 lần chạy trùng) ghi đè lẫn nhau.
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(CONFIG.SYNC_LOCK_WAIT_MS);
  } catch (lockErr) {
    acquired = false;
  }

  if (!acquired) {
    Logger.log('actionSync: khong lay duoc lock cho ' + xuong + ' ' + thang);
    return errorResponse('Đang có một phiên đồng bộ khác đang chạy. Vui lòng thử lại sau.', 'SYNC_BUSY');
  }

  try {
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
  } catch (err) {
    Logger.log('actionSync loi: ' + err + '\n' + err.stack);
    try {
      updateSyncStatus(xuong, thang, 0, 'LOI', String(err && err.message ? err.message : err));
    } catch (statusErr) {
      // không để lỗi khi ghi trạng thái che mất lỗi gốc
    }
    return errorResponse('Đồng bộ thất bại: ' + (err && err.message ? err.message : String(err)), 'SYNC_ERROR');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Thêm mới hoặc cập nhật nhân viên (dò theo MaNV). Nhân viên mới sẽ được đặt
 * mật khẩu ban đầu = Mã nhân viên, bắt buộc đổi mật khẩu lần đầu.
 *
 * Đọc toàn bộ vùng dữ liệu hiện có 1 lần (1 lệnh getValues), sửa/đè trong bộ
 * nhớ, rồi ghi lại toàn bộ vùng dữ liệu (cũ + mới) chỉ trong 1 lệnh setValues.
 */
function upsertEmployees(employees, xuong) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var data = (lastRow > 1) ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var byMaNV = {};
  data.forEach(function (row, i) {
    var maNV = String(row[idx['MaNV']] || '').trim();
    if (maNV) byMaNV[maNV] = i;
  });

  var now = formatDateVN(new Date());
  var count = 0;

  employees.forEach(function (e) {
    if (!e.maNV) return;
    var maNV = String(e.maNV).trim();
    var cccd = normalizeCCCD(e.cccd);
    count++;

    if (byMaNV.hasOwnProperty(maNV)) {
      var row = data[byMaNV[maNV]];
      // Cập nhật thông tin, KHÔNG đụng vào PasswordHash/Salt/MustChangePassword
      row[idx['CCCD']] = "'" + cccd; // giữ số 0 đầu
      row[idx['HoTen']] = e.hoTen || row[idx['HoTen']];
      row[idx['Xuong']] = xuong;
      row[idx['PhongBan']] = e.phongBan || '';
      row[idx['BoPhan']] = e.boPhan || '';
      row[idx['ChucVu']] = e.chucVu || '';
      row[idx['UpdatedAt']] = now;
    } else {
      var salt = generateSalt();
      var hash = hashPassword(maNV, salt); // mật khẩu ban đầu = mã nhân viên
      var newRow = new Array(headers.length).fill('');
      headers.forEach(function (h, i) {
        switch (h) {
          case 'CCCD': newRow[i] = "'" + cccd; break;
          case 'MaNV': newRow[i] = maNV; break;
          case 'HoTen': newRow[i] = e.hoTen || ''; break;
          case 'Xuong': newRow[i] = xuong; break;
          case 'PhongBan': newRow[i] = e.phongBan || ''; break;
          case 'BoPhan': newRow[i] = e.boPhan || ''; break;
          case 'ChucVu': newRow[i] = e.chucVu || ''; break;
          case 'PasswordHash': newRow[i] = hash; break;
          case 'PasswordSalt': newRow[i] = salt; break;
          case 'MustChangePassword': newRow[i] = true; break;
          case 'Role': newRow[i] = 'employee'; break;
          case 'UpdatedAt': newRow[i] = now; break;
          default: newRow[i] = '';
        }
      });
      byMaNV[maNV] = data.length;
      data.push(newRow);
    }
  });

  if (data.length > 0) {
    ensureSheetCapacity(sheet, data.length + 1, headers.length);
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

  return { count: count };
}

/**
 * Thêm mới hoặc cập nhật (upsert theo MaNV+Thang) phiếu lương. Cùng chiến
 * lược ghi theo lô như upsertEmployees ở trên.
 */
function upsertPayroll(payroll, xuong, thang) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var maNVCol = headers.indexOf('MaNV');
  var thangCol = headers.indexOf('Thang');

  var data = (lastRow > 1) ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var byKey = {};
  data.forEach(function (row, i) {
    var key = String(row[maNVCol] || '').trim() + '|' + row[thangCol];
    byKey[key] = i;
  });

  var now = formatDateVN(new Date());
  var count = 0;

  payroll.forEach(function (p) {
    if (!p.maNV) return;
    var key = String(p.maNV).trim() + '|' + thang;
    var rowData = buildPayrollRowData(p, xuong, thang, now, headers);

    if (byKey.hasOwnProperty(key)) {
      data[byKey[key]] = rowData;
    } else {
      byKey[key] = data.length;
      data.push(rowData);
    }
    count++;
  });

  if (data.length > 0) {
    ensureSheetCapacity(sheet, data.length + 1, headers.length);
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

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

function updateSyncStatus(xuong, thang, soDong, trangThai, ghiChu) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS,
    ['Xuong', 'Thang', 'LastSyncAt', 'SoDongDaXuLy', 'TrangThai', 'GhiChu']);
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 6);
  var data = (lastRow > 1) ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var now = formatDateVN(new Date());

  var foundIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === xuong && String(data[i][1]) === thang) {
      foundIndex = i;
      break;
    }
  }

  var rowValues = [xuong, thang, now, soDong, trangThai, ghiChu];
  if (foundIndex >= 0) {
    sheet.getRange(foundIndex + 2, 1, 1, 6).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}
