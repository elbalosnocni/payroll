function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function hashPassword(password, salt) {
  if (!CONFIG.PASSWORD_PEPPER) throw new Error('Chưa cấu hình PASSWORD_PEPPER.');
  var bytes = Utilities.newBlob(String(password) + '|' + salt + '|' + CONFIG.PASSWORD_PEPPER).getBytes();
  for (var i = 0; i < CONFIG.HASH_ITERATIONS; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return Utilities.base64Encode(bytes);
}
function verifyPassword(password, salt, expectedHash) {
  if (!salt || !expectedHash) return false;
  return safeCompare(hashPassword(password, salt), String(expectedHash));
}
function generateToken() { return Utilities.getUuid() + Utilities.getUuid(); }
function createSession(cccd, role) {
  var token = generateToken();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({ cccd: cccd, role: role, createdAt: Date.now() }), CONFIG.SESSION_DURATION_SEC);
  return token;
}
function getSession(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}
function destroySession(token) { if (token) CacheService.getScriptCache().remove('session_' + token); }

function findEmployeeByCCCD(cccd) {
  var normalized = normalizeCCCD(cccd);
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  requireHeaders(sheet, ['CCCD','MaNV','HoTen','PasswordHash','PasswordSalt','MustChangePassword','Role']);
  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) if (normalizeCCCD(rows[i].CCCD) === normalized) return rows[i];
  return null;
}

function actionLogin(params) {
  var cccd = normalizeCCCD(params.cccd);
  var password = String(params.password || '');
  if (!cccd || !password) return errorResponse('Vui lòng nhập đầy đủ Số căn cước và mật khẩu.', 'MISSING_FIELDS');
  var emp = findEmployeeByCCCD(cccd);
  if (!emp || !verifyPassword(password, emp.PasswordSalt, emp.PasswordHash)) {
    try { writeAudit(cccd, 'LOGIN_FAILED', cccd, 'Sai CCCD hoặc mật khẩu'); } catch (_) {}
    return errorResponse('Số căn cước hoặc mật khẩu không đúng.', 'INVALID_CREDENTIALS');
  }
  var role = String(emp.Role || '').toLowerCase() === 'admin' ? 'admin' : 'employee';
  var token = createSession(cccd, role);
  writeAudit(cccd, 'LOGIN_SUCCESS', cccd, 'role=' + role);
  return successResponse({ token: token, role: role, hoTen: emp.HoTen || '', mustChangePassword: isTrue(emp.MustChangePassword) });
}

function actionChangePassword(params) {
  var session = getSession(params.token);
  if (!session) return errorResponse('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.', 'SESSION_EXPIRED');
  var oldPassword = String(params.oldPassword || '');
  var newPassword = String(params.newPassword || '');
  var confirmPassword = String(params.confirmPassword || '');
  if (!oldPassword || !newPassword || !confirmPassword) return errorResponse('Vui lòng nhập đầy đủ thông tin.', 'MISSING_FIELDS');
  if (newPassword !== confirmPassword) return errorResponse('Mật khẩu mới và xác nhận mật khẩu không khớp.', 'PASSWORD_MISMATCH');
  if (newPassword.length < 6) return errorResponse('Mật khẩu mới phải có ít nhất 6 ký tự.', 'PASSWORD_TOO_SHORT');
  if (newPassword === oldPassword) return errorResponse('Mật khẩu mới phải khác mật khẩu hiện tại.', 'PASSWORD_SAME');

  var emp = findEmployeeByCCCD(session.cccd);
  if (!emp || !verifyPassword(oldPassword, emp.PasswordSalt, emp.PasswordHash)) {
    writeAudit(session.cccd, 'CHANGE_PASSWORD_FAILED', session.cccd, 'Sai mật khẩu hiện tại');
    return errorResponse('Mật khẩu hiện tại không đúng.', 'INVALID_OLD_PASSWORD');
  }

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var map = requireHeaders(sheet, ['PasswordHash','PasswordSalt','MustChangePassword','UpdatedAt']);
  var salt = generateSalt();
  sheet.getRange(emp.__row, map.PasswordHash).setValue(hashPassword(newPassword, salt));
  sheet.getRange(emp.__row, map.PasswordSalt).setValue(salt);
  sheet.getRange(emp.__row, map.MustChangePassword).setValue(false);
  sheet.getRange(emp.__row, map.UpdatedAt).setValue(formatDateVN(new Date()));
  writeAudit(session.cccd, 'CHANGE_PASSWORD_SUCCESS', session.cccd, '');
  return successResponse({ message: 'Đổi mật khẩu thành công.' });
}
function requireAdmin(token) { var s = getSession(token); return s && s.role === 'admin' ? s : null; }
function requireAuth(token) { return getSession(token); }
