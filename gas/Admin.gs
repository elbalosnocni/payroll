function actionAdminListEmployees(params) {
  var session = requireAdmin(params.token); if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');
  var rows = sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES));
  return successResponse({ employees: rows.map(function(r){ return { cccd:maskCCCD(r.CCCD), maNV:r.MaNV, hoTen:r.HoTen, xuong:r.Xuong, phongBan:r.PhongBan, boPhan:r.BoPhan, chucVu:r.ChucVu, role:String(r.Role||'employee'), mustChangePassword:isTrue(r.MustChangePassword), updatedAt:r.UpdatedAt||'' }; }) });
}
function actionAdminSearchEmployee(params) {
  var session = requireAdmin(params.token); if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');
  var key = String(params.keyword||'').toLowerCase().trim(); if (!key) return errorResponse('Vui lòng nhập từ khoá tìm kiếm.', 'MISSING_FIELDS');
  var keyC = normalizeCCCD(key);
  var rows = sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES)).filter(function(r){
    return String(r.HoTen||'').toLowerCase().indexOf(key)!==-1 || String(r.MaNV||'').toLowerCase().indexOf(key)!==-1 || (keyC && normalizeCCCD(r.CCCD).indexOf(keyC)!==-1);
  });
  return successResponse({results: rows.map(function(r){return {cccd:maskCCCD(r.CCCD),maNV:r.MaNV,hoTen:r.HoTen,xuong:r.Xuong,phongBan:r.PhongBan,boPhan:r.BoPhan,chucVu:r.ChucVu,role:String(r.Role||'employee'),mustChangePassword:isTrue(r.MustChangePassword)};})});
}
function actionAdminResetPassword(params) {
  var session = requireAdmin(params.token); if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');
  var maNV = String(params.maNV||'').trim(); if (!maNV) return errorResponse('Vui lòng chọn nhân viên cần reset.', 'MISSING_FIELDS');
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES), rows = sheetToObjects(sheet), target = rows.filter(function(r){return String(r.MaNV).trim()===maNV;})[0];
  if (!target) return errorResponse('Không tìm thấy nhân viên.', 'NOT_FOUND');
  var map = requireHeaders(sheet,['PasswordHash','PasswordSalt','MustChangePassword','UpdatedAt']);
  var salt = generateSalt();
  sheet.getRange(target.__row,map.PasswordHash).setValue(hashPassword(maNV,salt));
  sheet.getRange(target.__row,map.PasswordSalt).setValue(salt);
  sheet.getRange(target.__row,map.MustChangePassword).setValue(true);
  sheet.getRange(target.__row,map.UpdatedAt).setValue(formatDateVN(new Date()));
  writeAudit(session.cccd,'ADMIN_RESET_PASSWORD',maNV,'Reset về mã nhân viên');
  return successResponse({message:'Đã reset mật khẩu về Mã nhân viên cho '+(target.HoTen||maNV)+'.'});
}
function actionAdminSyncStatus(params) {
  var session=requireAdmin(params.token); if(!session) return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS));
  rows.sort(function(a,b){return String(b.LastSyncAt||'').localeCompare(String(a.LastSyncAt||''));});
  return successResponse({syncStatus:rows});
}
function actionAdminAuditLog(params) {
  var session=requireAdmin(params.token); if(!session) return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var limit=Math.max(1,Math.min(500,Number(params.limit)||200));
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_AUDIT)); rows.reverse();
  return successResponse({logs:rows.slice(0,limit)});
}
function maskCCCD(cccd){var s=normalizeCCCD(cccd); if(s.length<=4)return s; return '*'.repeat(Math.max(0,s.length-4))+s.slice(-4);}
