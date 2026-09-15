function actionLogin(p) {
  var cccd=normalizeCCCD(p.cccd), password=String(p.password||'');
  if(cccd.length<8 || !password) return errorResponse('Vui lòng nhập đầy đủ CCCD và mật khẩu.','MISSING_FIELDS');

  var throttleKey='LOGIN_'+cccd;
  var cache=CacheService.getScriptCache();
  if(cache.get(throttleKey)==='BLOCK') return errorResponse('Đăng nhập thất bại quá nhiều lần. Vui lòng thử lại sau 15 phút.','RATE_LIMIT');

  var emp=findEmployeeByCCCD(cccd);
  if(!emp) { recordLoginFail(throttleKey); return errorResponse('Sai CCCD hoặc mật khẩu.','INVALID_LOGIN'); }

  var valid=false;
  if(emp.PasswordHash && emp.PasswordSalt) valid=safeCompare(hashPassword(password,String(emp.PasswordSalt)),String(emp.PasswordHash));
  if(!valid) { recordLoginFail(throttleKey); return errorResponse('Sai CCCD hoặc mật khẩu.','INVALID_LOGIN'); }
  cache.remove(throttleKey);

  var token=newToken();
  var session={cccd:cccd,maNV:String(emp.MaNV),role:normalizeRole(emp.Role),mustChangePassword:isTrue(emp.MustChangePassword),createdAt:Date.now()};
  cache.put(tokenCacheKey(token),JSON.stringify(session),CONFIG.SESSION_TTL_SEC);
  writeAudit(cccd,'LOGIN',emp.MaNV,'Đăng nhập thành công');
  return successResponse({token:token,role:session.role,mustChangePassword:session.mustChangePassword,hoTen:emp.HoTen});
}
function recordLoginFail(key){
  var cache=CacheService.getScriptCache(), n=Number(cache.get(key)||0)+1;
  if(n>=CONFIG.MAX_LOGIN_ATTEMPTS) cache.put(key,'BLOCK',CONFIG.LOGIN_WINDOW_SEC);
  else cache.put(key,String(n),CONFIG.LOGIN_WINDOW_SEC);
}
function isTrue(v){ return v===true || String(v).toLowerCase()==='true' || String(v)==='1'; }
function normalizeRole(v){ return String(v||'EMPLOYEE').toUpperCase()==='ADMIN' ? 'ADMIN' : 'EMPLOYEE'; }

function requireAuth(token){
  if(!token) return null;
  var raw=CacheService.getScriptCache().get(tokenCacheKey(token));
  if(!raw) return null;
  try{return JSON.parse(raw);}catch(e){return null;}
}
function requireAdmin(token){
  var s=requireAuth(token); return s && s.role==='ADMIN' ? s : null;
}
function findEmployeeByCCCD(cccd){
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES));
  cccd=normalizeCCCD(cccd);
  return rows.filter(function(r){return normalizeCCCD(r.CCCD)===cccd;})[0] || null;
}
function findEmployeeByMaNV(maNV){
  var rows=sheetToObjects(getOrCreateSheet(CONFIG.SHEET_EMPLOYEES));
  return rows.filter(function(r){return String(r.MaNV).trim()===String(maNV).trim();})[0] || null;
}
function actionChangePassword(p){
  var s=requireAuth(p.token);
  if(!s) return errorResponse('Phiên đăng nhập đã hết hạn.','SESSION_EXPIRED');
  var oldPw=String(p.oldPassword||''), newPw=String(p.newPassword||''), confirm=String(p.confirmPassword||'');
  if(!oldPw||!newPw||!confirm) return errorResponse('Vui lòng nhập đủ 3 trường mật khẩu.','MISSING_FIELDS');
  if(newPw.length<8) return errorResponse('Mật khẩu mới phải có ít nhất 8 ký tự.','WEAK_PASSWORD');
  if(newPw!==confirm) return errorResponse('Xác nhận mật khẩu không khớp.','PASSWORD_MISMATCH');
  if(oldPw===newPw) return errorResponse('Mật khẩu mới phải khác mật khẩu hiện tại.','SAME_PASSWORD');

  var emp=findEmployeeByCCCD(s.cccd);
  if(!emp) return errorResponse('Không tìm thấy tài khoản.','NOT_FOUND');
  if(!safeCompare(hashPassword(oldPw,String(emp.PasswordSalt)),String(emp.PasswordHash))) return errorResponse('Mật khẩu hiện tại không đúng.','INVALID_PASSWORD');

  var salt=generateSalt(), hash=hashPassword(newPw,salt), sh=getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  setField(sh,emp.__row,'PasswordSalt',salt);
  setField(sh,emp.__row,'PasswordHash',hash);
  setField(sh,emp.__row,'MustChangePassword',false);
  setField(sh,emp.__row,'UpdatedAt',formatDateVN(new Date()));
  writeAudit(s.cccd,'CHANGE_PASSWORD',emp.MaNV,'Đổi mật khẩu thành công');

  // session mới, cập nhật cờ mustChangePassword
  var newToken=newToken(), ns={cccd:s.cccd,maNV:s.maNV,role:s.role,mustChangePassword:false,createdAt:Date.now()};
  CacheService.getScriptCache().put(tokenCacheKey(newToken),JSON.stringify(ns),CONFIG.SESSION_TTL_SEC);
  CacheService.getScriptCache().remove(tokenCacheKey(p.token));
  return successResponse({token:newToken,mustChangePassword:false,message:'Đổi mật khẩu thành công.'});
}
