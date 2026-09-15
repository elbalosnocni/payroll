function actionGetPayroll(p){
  var s=requireAuth(p.token);
  if(!s) return errorResponse('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.','SESSION_EXPIRED');
  var emp=findEmployeeByCCCD(s.cccd);
  if(!emp) return errorResponse('Không tìm thấy hồ sơ nhân viên.','NOT_FOUND');
  if(isTrue(emp.MustChangePassword)) return errorResponse('Bạn phải đổi mật khẩu trước khi xem phiếu lương.','MUST_CHANGE_PASSWORD');

  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_PAYROLL)).filter(function(r){return String(r.MaNV)===String(emp.MaNV);});
  if(!p.thang){
    return successResponse({hoTen:emp.HoTen,maNV:emp.MaNV,xuong:emp.Xuong,phongBan:emp.PhongBan,boPhan:emp.BoPhan,chucVu:emp.ChucVu,months:uniqueSortMonthsDesc(rows.map(function(r){return r.Thang;}))});
  }
  var rec=rows.filter(function(r){return String(r.Thang)===String(p.thang);})[0];
  if(!rec) return errorResponse('Không tìm thấy phiếu lương tháng '+p.thang+'.','NOT_FOUND');
  writeAudit(s.cccd,'VIEW_PAYROLL',emp.MaNV,'Tháng '+p.thang);
  return successResponse({
    thongTinNhanVien:{hoTen:emp.HoTen,maNV:emp.MaNV,xuong:emp.Xuong,phongBan:emp.PhongBan,boPhan:emp.BoPhan,chucVu:emp.ChucVu},
    thang:String(rec.Thang),updatedAt:String(rec.UpdatedAt||''),
    luongThang:{
      luongCoBan:formatCurrencyVND(rec.LuongCoBan),soNgayLamViec:formatDaysHours(rec.SoNgayLamViec),soNgayLe:formatDaysHours(rec.SoNgayLe),
      soNgayNghiHuongLuong:formatDaysHours(rec.SoNgayNghiHuongLuong),soNgayNghiKhongLuong:formatDaysHours(rec.SoNgayNghiKhongLuong),
      soNgayNghiHuongLuongToiThieuVung:formatDaysHours(rec.SoNgayNghiHuongLuongToiThieuVung),tongLuongThang:formatCurrencyVND(rec.LuongThang)
    },
    luongNgoaiGio:{
      soGioNgoaiGio:formatDaysHours(rec.SoGioNgoaiGio),soGioNgayNghi:formatDaysHours(rec.SoGioNgayNghi),
      soGioNgoaiGioNgayNghi:formatDaysHours(rec.SoGioNgoaiGioNgayNghi),soGioTangCaDem:formatDaysHours(rec.SoGioTangCaDem),
      soGioLamNgayLe:formatDaysHours(rec.SoGioLamNgayLe),soGioNgoaiGioNgayLe:formatDaysHours(rec.SoGioNgoaiGioNgayLe),
      soGioTangCaDemNgayLe:formatDaysHours(rec.SoGioTangCaDemNgayLe),soNgayLamCaDem:formatDaysHours(rec.SoNgayLamCaDem),
      tongLuongNgoaiGio:formatCurrencyVND(rec.LuongNgoaiGio)
    },
    phuCapTroCap:{
      tienKhac:formatCurrencyVND(rec.TienKhac),tienKyLuat:formatCurrencyVND(rec.TienKyLuat),tienGanBo2Nam:formatCurrencyVND(rec.TienGanBo2Nam),
      tienGanBo5Nam:formatCurrencyVND(rec.TienGanBo5Nam),tienGanBo10Nam:formatCurrencyVND(rec.TienGanBo10Nam),tienNhaO:formatCurrencyVND(rec.TienNhaO),
      tienDiLai:formatCurrencyVND(rec.TienDiLai),tienThuongChuyenCan:formatCurrencyVND(rec.TienThuongChuyenCan),
      hoaHongThuongVuotDinhMuc:formatCurrencyVND(rec.HoaHongThuongVuotDinhMuc),troCapThoiViecPhepNam:formatCurrencyVND(rec.TroCapThoiViecPhepNam)
    },
    tongThuNhap:formatCurrencyVND(rec.TongKhoanThuNhap),
    khauTru:{bhxh:formatCurrencyVND(rec.BHXH),bhyt:formatCurrencyVND(rec.BHYT),bhtn:formatCurrencyVND(rec.BHTN),
      thueThuNhap:formatCurrencyVND(rec.ThueThuNhap),tamUng:formatCurrencyVND(rec.TamUng),khauTruKhac:formatCurrencyVND(rec.KhauTruKhac)},
    luongThucLinh:formatCurrencyVND(rec.LuongThucLinh)
  });
}
function uniqueSortMonthsDesc(a){
  var u={}, out=[]; (a||[]).forEach(function(m){m=String(m||'');if(m&&!u[m]){u[m]=1;out.push(m);}});
  out.sort(function(x,y){var a=x.split('-'),b=y.split('-');return new Date(+b[1],+b[0]-1)-new Date(+a[1],+a[0]-1);}); return out;
}
