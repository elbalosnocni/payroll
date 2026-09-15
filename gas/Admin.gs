function actionAdminListEmployees(p){
  if(!requireAdmin(p.token))return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES));
  return successResponse({employees:rows.map(adminEmployeeView)});
}
function adminEmployeeView(r){return{cccd:maskCCCD(r.CCCD),maNV:String(r.MaNV),hoTen:String(r.HoTen||''),xuong:String(r.Xuong||''),phongBan:String(r.PhongBan||''),boPhan:String(r.BoPhan||''),chucVu:String(r.ChucVu||''),role:normalizeRole(r.Role),mustChangePassword:isTrue(r.MustChangePassword),updatedAt:String(r.UpdatedAt||'')};}
function actionAdminSearchEmployee(p){
  if(!requireAdmin(p.token))return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var k=normalizeText(p.keyword||''), kc=normalizeCCCD(p.keyword||''); if(!k&& !kc)return errorResponse('Vui lòng nhập từ khoá.','MISSING_FIELDS');
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES)).filter(function(r){
    return normalizeText(r.HoTen).indexOf(k)!==-1 || normalizeText(r.MaNV).indexOf(k)!==-1 || (kc&&normalizeCCCD(r.CCCD).indexOf(kc)!==-1);
  });
  return successResponse({results:rows.slice(0,100).map(adminEmployeeView)});
}
function actionAdminResetPassword(p){
  var s=requireAdmin(p.token); if(!s)return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var emp=findEmployeeByMaNV(p.maNV); if(!emp)return errorResponse('Không tìm thấy nhân viên.','NOT_FOUND');
  var salt=generateSalt(),sh=getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  setField(sh,emp.__row,'PasswordSalt',salt);setField(sh,emp.__row,'PasswordHash',hashPassword(emp.MaNV,salt));setField(sh,emp.__row,'MustChangePassword',true);setField(sh,emp.__row,'UpdatedAt',formatDateVN(new Date()));
  writeAudit(s.cccd,'ADMIN_RESET_PASSWORD',emp.MaNV,'Mật khẩu đặt lại về Mã NV');
  return successResponse({message:'Đã reset mật khẩu cho '+emp.HoTen+'. Mật khẩu mới là Mã NV và bắt buộc đổi khi đăng nhập.'});
}
function actionAdminSyncStatus(p){
  if(!requireAdmin(p.token))return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS));rows.sort(function(a,b){return parseVNDate(b.LastSyncAt)-parseVNDate(a.LastSyncAt);});
  return successResponse({syncStatus:rows});
}
function actionAdminAuditLog(p){
  if(!requireAdmin(p.token))return errorResponse('Bạn không có quyền truy cập.','FORBIDDEN');
  var limit=Math.min(Math.max(Number(p.limit)||200,1),500),rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_AUDIT));rows.reverse();
  return successResponse({logs:rows.slice(0,limit)});
}
function maskCCCD(v){var s=normalizeCCCD(v);return s.length<=4?s:'*'.repeat(s.length-4)+s.slice(-4);}
function parseVNDate(s){
  var m=String(s||'').match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
  return m?new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5],+m[6]):new Date(0);
}
