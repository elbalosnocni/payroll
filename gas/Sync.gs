function actionSync(p){
  if(!CONFIG.SYNC_API_KEY || !safeCompare(String(p.apiKey||''),CONFIG.SYNC_API_KEY)) return errorResponse('Sai API key.','FORBIDDEN');
  var xuong=String(p.xuong||'').trim(), thang=String(p.thang||'').trim();
  if(!/^(Snack|Flexible)$/i.test(xuong) || !/^(0[1-9]|1[0-2])-\d{4}$/.test(thang)) return errorResponse('Xưởng hoặc tháng không hợp lệ.','VALIDATION_ERROR');
  var lock=LockService.getScriptLock();
  try{lock.waitLock(30000);}catch(e){return errorResponse('Hệ thống đang đồng bộ dữ liệu khác. Thử lại sau.','BUSY');}
  try{
    var er=upsertEmployees(Array.isArray(p.employees)?p.employees:[],xuong);
    var pr=upsertPayroll(Array.isArray(p.payroll)?p.payroll:[],xuong,thang);
    updateSyncStatus(xuong,thang,pr.count,'OK',er.count+' nhân viên; '+pr.count+' phiếu lương');
    writeAudit('SYSTEM_VBA','SYNC',xuong+' '+thang,er.count+' employees, '+pr.count+' payroll');
    return successResponse({message:'Đồng bộ thành công.',employeesProcessed:er.count,payrollProcessed:pr.count});
  }catch(e){
    updateSyncStatus(xuong,thang,0,'ERROR',String(e.message||e));
    throw e;
  }finally{lock.releaseLock();}
}
function upsertEmployees(list,xuong){
  var sh=getOrCreateSheet(CONFIG.SHEET_EMPLOYEES), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var rows=sheetToObjects(sh), by={}; rows.forEach(function(r){by[String(r.MaNV).trim()]=r;});
  var now=formatDateVN(new Date()), count=0;
  (list||[]).forEach(function(e){
    var ma=String(e.maNV||'').trim(); if(!ma)return;
    var found=by[ma], cccd=normalizeCCCD(e.cccd);
    if(found){
      setField(sh,found.__row,'CCCD',cccd); setField(sh,found.__row,'HoTen',e.hoTen||found.HoTen||'');
      setField(sh,found.__row,'Xuong',xuong); setField(sh,found.__row,'PhongBan',e.phongBan||'');
      setField(sh,found.__row,'BoPhan',e.boPhan||''); setField(sh,found.__row,'ChucVu',e.chucVu||''); setField(sh,found.__row,'UpdatedAt',now);
      if(!found.PasswordHash || !found.PasswordSalt){
        var salt=generateSalt(); setField(sh,found.__row,'PasswordSalt',salt); setField(sh,found.__row,'PasswordHash',hashPassword(ma,salt));
        setField(sh,found.__row,'MustChangePassword',true);
      }
    }else{
      var salt2=generateSalt(), row=headers.map(function(h){
        switch(h){case'CCCD':return cccd;case'MaNV':return ma;case'HoTen':return e.hoTen||'';case'Xuong':return xuong;
        case'PhongBan':return e.phongBan||'';case'BoPhan':return e.boPhan||'';case'ChucVu':return e.chucVu||'';
        case'PasswordHash':return hashPassword(ma,salt2);case'PasswordSalt':return salt2;case'MustChangePassword':return true;
        case'Role':return'EMPLOYEE';case'UpdatedAt':return now;default:return'';}
      });
      sh.appendRow(row);
    } count++;
  });
  sh.getRange('A:A').setNumberFormat('@'); return{count:count};
}
function upsertPayroll(list,xuong,thang){
  var sh=getOrCreateSheet(CONFIG.SHEET_PAYROLL), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var rows=sheetToObjects(sh), by={}; rows.forEach(function(r){by[String(r.MaNV).trim()+'|'+String(r.Thang).trim()]=r;});
  var now=formatDateVN(new Date()), count=0;
  var batch=[];
  (list||[]).forEach(function(p){
    var ma=String(p.maNV||'').trim(); if(!ma)return;
    var row=buildPayrollRowData(p,xuong,thang,now,headers), found=by[ma+'|'+thang];
    if(found) sh.getRange(found.__row,1,1,headers.length).setValues([row]); else batch.push(row); count++;
  });
  if(batch.length) sh.getRange(sh.getLastRow()+1,1,batch.length,headers.length).setValues(batch);
  return{count:count};
}
function buildPayrollRowData(p,x,t,n,h){
  var m={MaNV:p.maNV||'',HoTen:p.hoTen||'',Xuong:x,Thang:t,LuongCoBan:p.luongCoBan,SoNgayLamViec:p.soNgayLamViec,SoNgayLe:p.soNgayLe,
  SoNgayNghiHuongLuong:p.soNgayNghiHuongLuong,SoNgayNghiKhongLuong:p.soNgayNghiKhongLuong,SoNgayNghiHuongLuongToiThieuVung:p.soNgayNghiHuongLuongToiThieuVung,
  LuongThang:p.luongThang,SoGioNgoaiGio:p.soGioNgoaiGio,SoGioNgayNghi:p.soGioNgayNghi,SoGioNgoaiGioNgayNghi:p.soGioNgoaiGioNgayNghi,
  SoGioTangCaDem:p.soGioTangCaDem,SoGioLamNgayLe:p.soGioLamNgayLe,SoGioNgoaiGioNgayLe:p.soGioNgoaiGioNgayLe,SoGioTangCaDemNgayLe:p.soGioTangCaDemNgayLe,
  SoNgayLamCaDem:p.soNgayLamCaDem,LuongNgoaiGio:p.luongNgoaiGio,TienKhac:p.tienKhac,TienKyLuat:p.tienKyLuat,TienGanBo2Nam:p.tienGanBo2Nam,
  TienGanBo5Nam:p.tienGanBo5Nam,TienGanBo10Nam:p.tienGanBo10Nam,TienNhaO:p.tienNhaO,TienDiLai:p.tienDiLai,TienThuongChuyenCan:p.tienThuongChuyenCan,
  HoaHongThuongVuotDinhMuc:p.hoaHongThuongVuotDinhMuc,TroCapThoiViecPhepNam:p.troCapThoiViecPhepNam,TongKhoanThuNhap:p.tongKhoanThuNhap,
  BHXH:p.bhxh,BHYT:p.bhyt,BHTN:p.bhtn,ThueThuNhap:p.thueThuNhap,TamUng:p.tamUng,KhauTruKhac:p.khauTruKhac,LuongThucLinh:p.luongThucLinh,UpdatedAt:n};
  return h.map(function(k){return m[k]!==undefined&&m[k]!==null&&m[k]!==''?m[k]:((k==='MaNV'||k==='HoTen'||k==='Xuong'||k==='Thang'||k==='UpdatedAt')?'':0);});
}
function updateSyncStatus(x,t,count,status,note){
  var sh=getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS), rows=sheetToObjects(sh), found=rows.filter(function(r){return String(r.Xuong)===x&&String(r.Thang)===t;})[0], now=formatDateVN(new Date());
  var vals=[x,t,now,count,status,note||'']; if(found)sh.getRange(found.__row,1,1,6).setValues([vals]); else sh.appendRow(vals);
}
