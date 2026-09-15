/**
 * Payroll.gs
 * -----------------------------------------------------------------------
 * Lấy phiếu lương của nhân viên đang đăng nhập. Dữ liệu được lọc CHẶT theo
 * MaNV gắn với session (không nhận MaNV/CCCD từ client) để nhân viên A không
 * thể sửa request để xem lương nhân viên B.
 * -----------------------------------------------------------------------
 */

/**
 * POST /getPayroll
 * body: { token, thang }   // thang là optional, ví dụ "08-2026". Nếu không
 *                          // truyền thì trả về danh sách các tháng có dữ liệu.
 */
function actionGetPayroll(params) {
  var session = requireAuth(params.token);
  if (!session) {
    return errorResponse('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.', 'SESSION_EXPIRED');
  }

  var emp = findEmployeeByCCCD(session.cccd);
  if (!emp) {
    return errorResponse('Không tìm thấy hồ sơ nhân viên.', 'NOT_FOUND');
  }

  var payrollSheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL);
  var allRows = sheetToObjects(payrollSheet);
  var myRows = allRows.filter(function (r) {
    return String(r.MaNV) === String(emp.MaNV);
  });

  if (!params.thang) {
    // Trả về danh sách các tháng có phiếu lương, mới nhất trước
    var months = myRows.map(function (r) { return r.Thang; });
    months = uniqueSortMonthsDesc(months);
    return successResponse({
      hoTen: emp.HoTen,
      maNV: emp.MaNV,
      xuong: emp.Xuong,
      phongBan: emp.PhongBan,
      boPhan: emp.BoPhan,
      chucVu: emp.ChucVu,
      months: months
    });
  }

  var record = myRows.filter(function (r) { return r.Thang === params.thang; })[0];
  if (!record) {
    return errorResponse('Không tìm thấy phiếu lương tháng ' + params.thang + '.', 'NOT_FOUND');
  }

  writeAudit(session.cccd, 'VIEW_PAYROLL', emp.MaNV, 'thang=' + params.thang);

  return successResponse({
    thongTinNhanVien: {
      hoTen: emp.HoTen,
      maNV: emp.MaNV,
      xuong: emp.Xuong,
      phongBan: emp.PhongBan,
      boPhan: emp.BoPhan,
      chucVu: emp.ChucVu
    },
    thang: record.Thang,
    updatedAt: record.UpdatedAt,
    luongThang: {
      luongCoBan: formatCurrencyVND(record.LuongCoBan),
      soNgayLamViec: formatDaysHours(record.SoNgayLamViec),
      soNgayLe: formatDaysHours(record.SoNgayLe),
      soNgayNghiHuongLuong: formatDaysHours(record.SoNgayNghiHuongLuong),
      soNgayNghiKhongLuong: formatDaysHours(record.SoNgayNghiKhongLuong),
      soNgayNghiHuongLuongToiThieuVung: formatDaysHours(record.SoNgayNghiHuongLuongToiThieuVung),
      tongLuongThang: formatCurrencyVND(record.LuongThang)
    },
    luongNgoaiGio: {
      soGioNgoaiGio: formatDaysHours(record.SoGioNgoaiGio),
      soGioNgayNghi: formatDaysHours(record.SoGioNgayNghi),
      soGioNgoaiGioNgayNghi: formatDaysHours(record.SoGioNgoaiGioNgayNghi),
      soGioTangCaDem: formatDaysHours(record.SoGioTangCaDem),
      soGioLamNgayLe: formatDaysHours(record.SoGioLamNgayLe),
      soGioNgoaiGioNgayLe: formatDaysHours(record.SoGioNgoaiGioNgayLe),
      soGioTangCaDemNgayLe: formatDaysHours(record.SoGioTangCaDemNgayLe),
      soNgayLamCaDem: formatDaysHours(record.SoNgayLamCaDem),
      tongLuongNgoaiGio: formatCurrencyVND(record.LuongNgoaiGio)
    },
    phuCapTroCap: {
      tienKhac: formatCurrencyVND(record.TienKhac),
      tienKyLuat: formatCurrencyVND(record.TienKyLuat),
      tienGanBo2Nam: formatCurrencyVND(record.TienGanBo2Nam),
      tienGanBo5Nam: formatCurrencyVND(record.TienGanBo5Nam),
      tienGanBo10Nam: formatCurrencyVND(record.TienGanBo10Nam),
      tienNhaO: formatCurrencyVND(record.TienNhaO),
      tienDiLai: formatCurrencyVND(record.TienDiLai),
      tienThuongChuyenCan: formatCurrencyVND(record.TienThuongChuyenCan),
      hoaHongThuongVuotDinhMuc: formatCurrencyVND(record.HoaHongThuongVuotDinhMuc),
      troCapThoiViecPhepNam: formatCurrencyVND(record.TroCapThoiViecPhepNam)
    },
    tongThuNhap: formatCurrencyVND(record.TongKhoanThuNhap),
    khauTru: {
      bhxh: formatCurrencyVND(record.BHXH),
      bhyt: formatCurrencyVND(record.BHYT),
      bhtn: formatCurrencyVND(record.BHTN),
      thueThuNhap: formatCurrencyVND(record.ThueThuNhap),
      tamUng: formatCurrencyVND(record.TamUng),
      khauTruKhac: formatCurrencyVND(record.KhauTruKhac)
    },
    luongThucLinh: formatCurrencyVND(record.LuongThucLinh)
  });
}

/**
 * Sắp xếp danh sách tháng "MM-YYYY" giảm dần, loại trùng.
 */
function uniqueSortMonthsDesc(months) {
  var seen = {};
  var unique = [];
  months.forEach(function (m) {
    if (m && !seen[m]) {
      seen[m] = true;
      unique.push(m);
    }
  });
  unique.sort(function (a, b) {
    var pa = a.split('-'), pb = b.split('-');
    var da = new Date(parseInt(pa[1], 10), parseInt(pa[0], 10) - 1, 1);
    var db = new Date(parseInt(pb[1], 10), parseInt(pb[0], 10) - 1, 1);
    return db - da;
  });
  return unique;
}
