/**
 * Sync.gs
 *
 * Nhận dữ liệu lương từ VBA.
 *
 * Điểm quan trọng:
 *  - VBA phải gửi action = "sync".
 *  - Employees upsert theo MaNV.
 *  - Payroll upsert theo MaNV + Thang.
 *  - Ghi dữ liệu theo BATCH (một lần setValues), tránh appendRow/setValue
 *    hàng trăm lần làm request VBA bị timeout.
 */

function actionSync(params) {
  if (!params.apiKey || !safeCompare(String(params.apiKey), CONFIG.SYNC_API_KEY)) {
    return errorResponse('Sai API key.', 'FORBIDDEN');
  }

  var xuong = String(params.xuong || '').trim();
  var thang = String(params.thang || '').trim();
  if (!xuong || !thang) {
    return errorResponse('Thieu thong tin xuong/thang.', 'MISSING_FIELDS');
  }

  var employees = Array.isArray(params.employees) ? params.employees : [];
  var payroll = Array.isArray(params.payroll) ? params.payroll : [];

  if (employees.length === 0 && payroll.length === 0) {
    return errorResponse('Khong co du lieu employees/payroll.', 'EMPTY_DATA');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return errorResponse('Dang co mot phien dong bo khac dang chay. Vui long thu lai sau.', 'SYNC_BUSY');
  }

  try {
    var empResult = upsertEmployeesBatch(employees, xuong);
    var payrollResult = upsertPayrollBatch(payroll, xuong, thang);

    var total = empResult.count + payrollResult.count;

    updateSyncStatus(
      xuong,
      thang,
      total,
      'OK',
      empResult.count + ' nhan vien, ' + payrollResult.count + ' phieu luong'
    );

    writeAudit(
      'SYSTEM_VBA',
      'SYNC',
      xuong + ' ' + thang,
      empResult.count + ' employees, ' + payrollResult.count + ' payroll rows'
    );

    return successResponse({
      message: 'Dong bo thanh cong.',
      employeesProcessed: empResult.count,
      payrollProcessed: payrollResult.count
    });
  } catch (err) {
    Logger.log('actionSync error: ' + err + '\n' + err.stack);
    return errorResponse('Dong bo that bai: ' + err.message, 'SYNC_ERROR');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Upsert Employees bang batch.
 * Khong dung appendRow/setValue cho tung nhan vien.
 */
function upsertEmployeesBatch(employees, xuong) {
  var defaultHeaders = [
    'CCCD', 'MaNV', 'HoTen', 'Xuong', 'PhongBan', 'BoPhan', 'ChucVu',
    'PasswordHash', 'PasswordSalt', 'MustChangePassword', 'Role', 'UpdatedAt'
  ];

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES, defaultHeaders);
  ensureHeaders(sheet, defaultHeaders);

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getDataRange().getValues();
  var headerIndex = makeHeaderIndex(headers);

  var maNvIndex = headerIndex.MaNV;
  if (maNvIndex === undefined) {
    throw new Error('Sheet Employees thieu cot MaNV.');
  }

  var rowByMaNV = {};
  for (var i = 1; i < values.length; i++) {
    var existingMaNV = String(values[i][maNvIndex] || '').trim();
    if (existingMaNV) rowByMaNV[existingMaNV] = i;
  }

  var now = formatDateVN(new Date());
  var count = 0;
  var changed = false;

  employees.forEach(function (e) {
    if (!e || !e.maNV) return;

    var maNV = String(e.maNV).trim();
    if (!maNV) return;

    var rowIndex = rowByMaNV[maNV];
    var row;

    if (rowIndex !== undefined) {
      row = values[rowIndex];

      setIfColumn(row, headerIndex, 'CCCD', "'" + normalizeCCCD(e.cccd));
      setIfColumn(row, headerIndex, 'MaNV', maNV);
      setIfColumnIfNotBlank(row, headerIndex, 'HoTen', e.hoTen);
      setIfColumn(row, headerIndex, 'Xuong', xuong);
      setIfColumn(row, headerIndex, 'PhongBan', e.phongBan || '');
      setIfColumn(row, headerIndex, 'BoPhan', e.boPhan || '');
      setIfColumn(row, headerIndex, 'ChucVu', e.chucVu || '');
      setIfColumn(row, headerIndex, 'UpdatedAt', now);
      changed = true;
    } else {
      row = new Array(headers.length).fill('');
      var salt = generateSalt();
      var hash = hashPassword(maNV, salt);

      setIfColumn(row, headerIndex, 'CCCD', "'" + normalizeCCCD(e.cccd));
      setIfColumn(row, headerIndex, 'MaNV', maNV);
      setIfColumn(row, headerIndex, 'HoTen', e.hoTen || '');
      setIfColumn(row, headerIndex, 'Xuong', xuong);
      setIfColumn(row, headerIndex, 'PhongBan', e.phongBan || '');
      setIfColumn(row, headerIndex, 'BoPhan', e.boPhan || '');
      setIfColumn(row, headerIndex, 'ChucVu', e.chucVu || '');
      setIfColumn(row, headerIndex, 'PasswordHash', hash);
      setIfColumn(row, headerIndex, 'PasswordSalt', salt);
      setIfColumn(row, headerIndex, 'MustChangePassword', true);
      setIfColumn(row, headerIndex, 'Role', 'employee');
      setIfColumn(row, headerIndex, 'UpdatedAt', now);

      values.push(row);
      rowByMaNV[maNV] = values.length - 1;
      changed = true;
    }

    count++;
  });

  if (changed && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));
  }

  // Dinh dang CCCD la text de giu so 0 dau.
  if (headerIndex.CCCD !== undefined && sheet.getLastRow() >= 2) {
    sheet.getRange(2, headerIndex.CCCD + 1, sheet.getLastRow() - 1, 1).setNumberFormat('@');
  }

  return { count: count };
}

/**
 * Upsert Payroll bang batch theo MaNV + Thang.
 */
function upsertPayrollBatch(payroll, xuong, thang) {
  var defaultHeaders = [
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
  ];

  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL, defaultHeaders);
  ensureHeaders(sheet, defaultHeaders);

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getDataRange().getValues();
  var headerIndex = makeHeaderIndex(headers);

  if (headerIndex.MaNV === undefined || headerIndex.Thang === undefined) {
    throw new Error('Sheet Payroll thieu cot MaNV hoac Thang.');
  }

  var rowByKey = {};
  for (var i = 1; i < values.length; i++) {
    var existingMaNV = String(values[i][headerIndex.MaNV] || '').trim();
    var existingThang = String(values[i][headerIndex.Thang] || '').trim();
    if (existingMaNV && existingThang) {
      rowByKey[existingMaNV + '|' + existingThang] = i;
    }
  }

  var now = formatDateVN(new Date());
  var count = 0;
  var changed = false;

  payroll.forEach(function (p) {
    if (!p || !p.maNV) return;

    var maNV = String(p.maNV).trim();
    if (!maNV) return;

    var key = maNV + '|' + thang;
    var rowIndex = rowByKey[key];
    var row = buildPayrollRowDataBatch(p, xuong, thang, now, headers, headerIndex);

    if (rowIndex !== undefined) {
      values[rowIndex] = row;
    } else {
      values.push(row);
      rowByKey[key] = values.length - 1;
    }

    changed = true;
    count++;
  });

  if (changed && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));
  }

  return { count: count };
}

function buildPayrollRowDataBatch(p, xuong, thang, now, headers, headerIndex) {
  var map = {
    MaNV: p.maNV,
    HoTen: p.hoTen || '',
    Xuong: xuong,
    Thang: thang,
    LuongCoBan: numberOrZero(p.luongCoBan),
    SoNgayLamViec: numberOrZero(p.soNgayLamViec),
    SoNgayLe: numberOrZero(p.soNgayLe),
    SoNgayNghiHuongLuong: numberOrZero(p.soNgayNghiHuongLuong),
    SoNgayNghiKhongLuong: numberOrZero(p.soNgayNghiKhongLuong),
    SoNgayNghiHuongLuongToiThieuVung: numberOrZero(p.soNgayNghiHuongLuongToiThieuVung),
    LuongThang: numberOrZero(p.luongThang),
    SoGioNgoaiGio: numberOrZero(p.soGioNgoaiGio),
    SoGioNgayNghi: numberOrZero(p.soGioNgayNghi),
    SoGioNgoaiGioNgayNghi: numberOrZero(p.soGioNgoaiGioNgayNghi),
    SoGioTangCaDem: numberOrZero(p.soGioTangCaDem),
    SoGioLamNgayLe: numberOrZero(p.soGioLamNgayLe),
    SoGioNgoaiGioNgayLe: numberOrZero(p.soGioNgoaiGioNgayLe),
    SoGioTangCaDemNgayLe: numberOrZero(p.soGioTangCaDemNgayLe),
    SoNgayLamCaDem: numberOrZero(p.soNgayLamCaDem),
    LuongNgoaiGio: numberOrZero(p.luongNgoaiGio),
    TienKhac: numberOrZero(p.tienKhac),
    TienKyLuat: numberOrZero(p.tienKyLuat),
    TienGanBo2Nam: numberOrZero(p.tienGanBo2Nam),
    TienGanBo5Nam: numberOrZero(p.tienGanBo5Nam),
    TienGanBo10Nam: numberOrZero(p.tienGanBo10Nam),
    TienNhaO: numberOrZero(p.tienNhaO),
    TienDiLai: numberOrZero(p.tienDiLai),
    TienThuongChuyenCan: numberOrZero(p.tienThuongChuyenCan),
    HoaHongThuongVuotDinhMuc: numberOrZero(p.hoaHongThuongVuotDinhMuc),
    TroCapThoiViecPhepNam: numberOrZero(p.troCapThoiViecPhepNam),
    TongKhoanThuNhap: numberOrZero(p.tongKhoanThuNhap),
    BHXH: numberOrZero(p.bhxh),
    BHYT: numberOrZero(p.bhyt),
    BHTN: numberOrZero(p.bhtn),
    ThueThuNhap: numberOrZero(p.thueThuNhap),
    TamUng: numberOrZero(p.tamUng),
    KhauTruKhac: numberOrZero(p.khauTruKhac),
    LuongThucLinh: numberOrZero(p.luongThucLinh),
    UpdatedAt: now
  };

  return headers.map(function (h) {
    return Object.prototype.hasOwnProperty.call(map, h) ? map[h] : '';
  });
}

function numberOrZero(value) {
  var n = Number(value);
  return isFinite(n) ? n : 0;
}

function makeHeaderIndex(headers) {
  var result = {};
  headers.forEach(function (h, i) {
    result[String(h).trim()] = i;
  });
  return result;
}

function setIfColumn(row, headerIndex, name, value) {
  if (headerIndex[name] !== undefined) row[headerIndex[name]] = value;
}

function setIfColumnIfNotBlank(row, headerIndex, name, value) {
  if (headerIndex[name] !== undefined && value !== undefined && value !== null && String(value).trim() !== '') {
    row[headerIndex[name]] = value;
  }
}

function ensureHeaders(sheet, defaultHeaders) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.every(function (h) { return String(h || '').trim() === ''; })) {
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    sheet.setFrozenRows(1);
  }
}

function updateSyncStatus(xuong, thang, soDong, trangThai, ghiChu) {
  var sheet = getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS, [
    'Xuong', 'Thang', 'LastSyncAt', 'SoDongDaXuLy', 'TrangThai', 'GhiChu'
  ]);
  ensureHeaders(sheet, [
    'Xuong', 'Thang', 'LastSyncAt', 'SoDongDaXuLy', 'TrangThai', 'GhiChu'
  ]);

  var rows = sheetToObjects(sheet);
  var found = rows.filter(function (r) {
    return String(r.Xuong) === String(xuong) && String(r.Thang) === String(thang);
  })[0];

  var now = formatDateVN(new Date());
  if (found) {
    sheet.getRange(found.__row, 1, 1, 6).setValues([[xuong, thang, now, soDong, trangThai, ghiChu]]);
  } else {
    sheet.appendRow([xuong, thang, now, soDong, trangThai, ghiChu]);
  }
}
