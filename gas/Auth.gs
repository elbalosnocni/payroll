/**
 * Auth.gs
 * -----------------------------------------------------------------------
 * Đăng nhập, đổi mật khẩu, quản lý session (token), phân quyền employee/admin.
 *
 * Lưu ý bảo mật:
 * - Mật khẩu KHÔNG BAO GIỜ được lưu ở dạng plaintext.
 * - Băm mật khẩu bằng SHA-256 lặp nhiều vòng (HASH_ITERATIONS) kèm salt riêng
 *   cho từng user + pepper chung (PASSWORD_PEPPER) => tương đương PBKDF2 đơn giản.
 * - Session token là chuỗi ngẫu nhiên, lưu trong CacheService (không lưu trong
 *   Sheet để tránh lộ khi share quyền xem Sheet), có thời hạn SESSION_DURATION_SEC.
 * -----------------------------------------------------------------------
 */

/**
 * Sinh salt ngẫu nhiên (hex, 32 ký tự).
 */
function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

/**
 * Băm mật khẩu: lặp SHA-256 nhiều vòng với salt + pepper.
 */
function hashPassword(password, salt) {
  var input = String(password) + '|' + salt + '|' + CONFIG.PASSWORD_PEPPER;
  var digestBytes = Utilities.newBlob(input).getBytes();
  for (var i = 0; i < CONFIG.HASH_ITERATIONS; i++) {
    digestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digestBytes);
  }
  return Utilities.base64Encode(digestBytes);
}

function verifyPassword(password, salt, expectedHash) {
  var actualHash = hashPassword(password, salt);
  return safeCompare(actualHash, expectedHash);
}

/**
 * Sinh token session ngẫu nhiên.
 */
function generateToken() {
  return Utilities.getUuid() + Utilities.getUuid();
}

/**
 * Tạo session mới cho 1 user (employee hoặc admin), lưu vào CacheService.
 * Trả về token.
 */
function createSession(cccd, role) {
  var token = generateToken();
  var cache = CacheService.getScriptCache();
  var payload = JSON.stringify({ cccd: cccd, role: role, createdAt: new Date().getTime() });
  cache.put('session_' + token, payload, CONFIG.SESSION_DURATION_SEC);
  return token;
}

/**
 * Lấy thông tin session từ token. Trả về null nếu không hợp lệ / hết hạn.
 */
function getSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var raw = cache.get('session_' + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function destroySession(token) {
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + token);
}

/**
 * Tìm dòng nhân viên theo CCCD trong sheet Employees.
 */
function findEmployeeByCCCD(cccd) {
  var normalized = normalizeCCCD(cccd);
  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (normalizeCCCD(rows[i].CCCD) === normalized) {
      return rows[i];
    }
  }
  return null;
}

/**
 * POST /login
 * body: { cccd, password }
 */
function actionLogin(params) {
  var cccd = normalizeCCCD(params.cccd);
  var password = String(params.password || '');

  if (!cccd || !password) {
    return errorResponse('Vui lòng nhập đầy đủ Số căn cước và mật khẩu.', 'MISSING_FIELDS');
  }

  var emp = findEmployeeByCCCD(cccd);
  if (!emp) {
    writeAudit(cccd, 'LOGIN_FAILED', cccd, 'Không tìm thấy CCCD');
    return errorResponse('Số căn cước hoặc mật khẩu không đúng.', 'INVALID_CREDENTIALS');
  }

  if (!emp.PasswordHash || !emp.PasswordSalt) {
    writeAudit(cccd, 'LOGIN_FAILED', cccd, 'Tài khoản chưa có mật khẩu, cần admin reset');
    return errorResponse('Tài khoản chưa được khởi tạo mật khẩu. Vui lòng liên hệ Admin.', 'NO_PASSWORD');
  }

  var valid = verifyPassword(password, emp.PasswordSalt, emp.PasswordHash);
  if (!valid) {
    writeAudit(cccd, 'LOGIN_FAILED', cccd, 'Sai mật khẩu');
    return errorResponse('Số căn cước hoặc mật khẩu không đúng.', 'INVALID_CREDENTIALS');
  }

  var role = emp.Role === 'admin' ? 'admin' : 'employee';
  var token = createSession(cccd, role);
  writeAudit(cccd, 'LOGIN_SUCCESS', cccd, 'role=' + role);

  return successResponse({
    token: token,
    role: role,
    hoTen: emp.HoTen,
    mustChangePassword: emp.MustChangePassword === true || emp.MustChangePassword === 'TRUE' || emp.MustChangePassword === 'true'
  });
}

/**
 * POST /changePassword
 * body: { token, oldPassword, newPassword, confirmPassword }
 */
function actionChangePassword(params) {
  var session = getSession(params.token);
  if (!session) {
    return errorResponse('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.', 'SESSION_EXPIRED');
  }

  var oldPassword = String(params.oldPassword || '');
  var newPassword = String(params.newPassword || '');
  var confirmPassword = String(params.confirmPassword || '');

  if (!oldPassword || !newPassword || !confirmPassword) {
    return errorResponse('Vui lòng nhập đầy đủ thông tin.', 'MISSING_FIELDS');
  }
  if (newPassword !== confirmPassword) {
    return errorResponse('Mật khẩu mới và xác nhận mật khẩu không khớp.', 'PASSWORD_MISMATCH');
  }
  if (newPassword.length < 6) {
    return errorResponse('Mật khẩu mới phải có ít nhất 6 ký tự.', 'PASSWORD_TOO_SHORT');
  }
  if (newPassword === oldPassword) {
    return errorResponse('Mật khẩu mới phải khác mật khẩu hiện tại.', 'PASSWORD_SAME');
  }

  var emp = findEmployeeByCCCD(session.cccd);
  if (!emp) {
    return errorResponse('Không tìm thấy tài khoản.', 'NOT_FOUND');
  }

  var validOld = verifyPassword(oldPassword, emp.PasswordSalt, emp.PasswordHash);
  if (!validOld) {
    writeAudit(session.cccd, 'CHANGE_PASSWORD_FAILED', session.cccd, 'Sai mật khẩu hiện tại');
    return errorResponse('Mật khẩu hiện tại không đúng.', 'INVALID_OLD_PASSWORD');
  }

  var newSalt = generateSalt();
  var newHash = hashPassword(newPassword, newSalt);

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colHash = headers.indexOf('PasswordHash') + 1;
  var colSalt = headers.indexOf('PasswordSalt') + 1;
  var colMustChange = headers.indexOf('MustChangePassword') + 1;

  sheet.getRange(emp.__row, colHash).setValue(newHash);
  sheet.getRange(emp.__row, colSalt).setValue(newSalt);
  sheet.getRange(emp.__row, colMustChange).setValue(false);

  writeAudit(session.cccd, 'CHANGE_PASSWORD_SUCCESS', session.cccd, '');

  return successResponse({ message: 'Đổi mật khẩu thành công.' });
}

/**
 * Kiểm tra session hợp lệ và bắt buộc phải là admin. Trả về session object
 * nếu hợp lệ, hoặc null nếu không.
 */
function requireAdmin(token) {
  var session = getSession(token);
  if (!session || session.role !== 'admin') return null;
  return session;
}

function requireAuth(token) {
  return getSession(token);
}
