function actionGetPayroll(params) {
  var session = requireAuth(params.token);
  if (!session) return errorResponse('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.', 'SESSION_EXPIRED');
  var emp = findEmployeeByCCCD(session.cccd);
  if (!emp) return errorResponse('Không tìm thấy hồ sơ nhân viên.', 'NOT_FOUND');

  var sheet = getOrCreateSheet(CONFIG.SHEET_PAYROLL);
  requireHeaders(sheet, ['MaNV','Thang','LuongThucLinh','UpdatedAt']);
  var rows = sheetToObjects(sheet).filter(function(r){ return String(r.MaNV).trim() === String(emp.MaNV).trim(); });

  if (!params.thang) {
    return successResponse({
      hoTen: emp.HoTen || '', maNV: emp.MaNV || '', xuong: emp.Xuong || '',
      phongBan: emp.PhongBan || '', boPhan: emp.BoPhan || '', chucVu: emp.ChucVu || '',
      months: uniqueSortMonthsDesc(rows.map(function(r){ return String(r.Thang || '').trim(); }))
    });
  }

  var thang = String(params.thang).trim();
  var record = rows.filter(function(r){ return String(r.Thang).trim() === thang; })[0];
  if (!record) return errorResponse('Không tìm thấy phiếu lương tháng ' + thang + '.', 'NOT_FOUND');
  writeAudit(session.cccd, 'VIEW_PAYROLL', emp.MaNV, 'thang=' + thang);

  return successResponse({
    thongTinNhanVien: { hoTen:emp.HoTen||'', maNV:emp.MaNV||'', xuong:emp.Xuong||'', phongBan:emp.PhongBan||'', boPhan:emp.BoPhan||'', chucVu:emp.ChucVu||'' },
    thang: thang, updatedAt: record.UpdatedAt || '',
    luongThang: {
      luongCoBan:formatCurrencyVND(record.LuongCoBan), soNgayLamViec:formatDaysHours(record.SoNgayLamViec),
      soNgayLe:formatDaysHours(record.SoNgayLe), soNgayNghiHuongLuong:formatDaysHours(record.SoNgayNghiHuongLuong),
      soNgayNghiKhongLuong:formatDaysHours(record.SoNgayNghiKhongLuong),
      soNgayNghiHuongLuongToiThieuVung:formatDaysHours(record.SoNgayNghiHuongLuongToiThieuVung),
      tongLuongThang:formatCurrencyVND(record.LuongThang)
    },
    luongNgoaiGio: {
      soGioNgoaiGio:formatDaysHours(record.SoGioNgoaiGio), soGioNgayNghi:formatDaysHours(record.SoGioNgayNghi),
      soGioNgoaiGioNgayNghi:formatDaysHours(record.SoGioNgoaiGioNgayNghi), soGioTangCaDem:formatDaysHours(record.SoGioTangCaDem),
      soGioLamNgayLe:formatDaysHours(record.SoGioLamNgayLe), soGioNgoaiGioNgayLe:formatDaysHours(record.SoGioNgoaiGioNgayLe),
      soGioTangCaDemNgayLe:formatDaysHours(record.SoGioTangCaDemNgayLe), soNgayLamCaDem:formatDaysHours(record.SoNgayLamCaDem),
      tongLuongNgoaiGio:formatCurrencyVND(record.LuongNgoaiGio)
    },
    phuCapTroCap: {
      tienKhac:formatCurrencyVND(record.TienKhac), tienKyLuat:formatCurrencyVND(record.TienKyLuat),
      tienGanBo2Nam:formatCurrencyVND(record.TienGanBo2Nam), tienGanBo5Nam:formatCurrencyVND(record.TienGanBo5Nam),
      tienGanBo10Nam:formatCurrencyVND(record.TienGanBo10Nam), tienNhaO:formatCurrencyVND(record.TienNhaO),
      tienDiLai:formatCurrencyVND(record.TienDiLai), tienThuongChuyenCan:formatCurrencyVND(record.TienThuongChuyenCan),
      hoaHongThuongVuotDinhMuc:formatCurrencyVND(record.HoaHongThuongVuotDinhMuc),
      troCapThoiViecPhepNam:formatCurrencyVND(record.TroCapThoiViecPhepNam)
    },
    tongThuNhap:formatCurrencyVND(record.TongKhoanThuNhap),
    khauTru: {
      bhxh:formatCurrencyVND(record.BHXH), bhyt:formatCurrencyVND(record.BHYT), bhtn:formatCurrencyVND(record.BHTN),
      thueThuNhap:formatCurrencyVND(record.ThueThuNhap), tamUng:formatCurrencyVND(record.TamUng), khauTruKhac:formatCurrencyVND(record.KhauTruKhac)
    },
    luongThucLinh:formatCurrencyVND(record.LuongThucLinh)
  });
}

function uniqueSortMonthsDesc(months) {
  var seen = {}, out = [];
  months.forEach(function(m){ if (/^\d{2}-\d{4}$/.test(m) && !seen[m]) { seen[m]=true; out.push(m); } });
  out.sort(function(a,b){ var pa=a.split('-'), pb=b.split('-'); return (new Date(+pb[1],+pb[0]-1,1)) - (new Date(+pa[1],+pa[0]-1,1)); });
  return out;
}
