/**
 * Admin.gs
 * -----------------------------------------------------------------------
 * Toàn bộ hàm ở đây yêu cầu session role = 'admin'. Để tạo tài khoản admin
 * đầu tiên: vào sheet Employees, thêm 1 dòng với Role = admin và tự chạy hàm
 * adminSetInitialPassword() bên dưới trong Apps Script Editor (hoặc set
 * MustChangePassword = true + PasswordHash/Salt rỗng rồi dùng Reset Password).
 * -----------------------------------------------------------------------
 */

/**
 * POST /adminListEmployees
 * body: { token }
 */
function actionAdminListEmployees(params) {
  var session = requireAdmin(params.token);
  if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var rows = sheetToObjects(sheet);
  var list = rows.map(function (r) {
    return {
      cccd: maskCCCD(r.CCCD),
      maNV: r.MaNV,
      hoTen: r.HoTen,
      xuong: r.Xuong,
      phongBan: r.PhongBan,
      boPhan: r.BoPhan,
      chucVu: r.ChucVu,
      role: r.Role || 'employee',
      mustChangePassword: !!r.MustChangePassword,
      updatedAt: r.UpdatedAt
    };
  });
  return successResponse({ employees: list });
}

/**
 * POST /adminSearchEmployee
 * body: { token, keyword }  // tìm theo tên, mã NV, hoặc CCCD (đầy đủ)
 */
function actionAdminSearchEmployee(params) {
  var session = requireAdmin(params.token);
  if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');

  var keyword = String(params.keyword || '').toLowerCase().trim();
  if (!keyword) return errorResponse('Vui lòng nhập từ khoá tìm kiếm.', 'MISSING_FIELDS');

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var rows = sheetToObjects(sheet);
  var results = rows.filter(function (r) {
    return String(r.HoTen).toLowerCase().indexOf(keyword) !== -1 ||
      String(r.MaNV).toLowerCase().indexOf(keyword) !== -1 ||
      normalizeCCCD(r.CCCD).indexOf(normalizeCCCD(keyword)) !== -1;
  }).map(function (r) {
    return {
      cccd: maskCCCD(r.CCCD),
      maNV: r.MaNV,
      hoTen: r.HoTen,
      xuong: r.Xuong,
      phongBan: r.PhongBan,
      boPhan: r.BoPhan,
      chucVu: r.ChucVu,
      role: r.Role || 'employee',
      mustChangePassword: !!r.MustChangePassword
    };
  });

  return successResponse({ results: results });
}

/**
 * POST /adminResetPassword
 * body: { token, maNV }
 * Đặt lại mật khẩu = Mã nhân viên, bắt buộc đổi lại ở lần đăng nhập tới.
 */
function actionAdminResetPassword(params) {
  var session = requireAdmin(params.token);
  if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');

  var maNV = String(params.maNV || '').trim();
  if (!maNV) return errorResponse('Vui lòng chọn nhân viên cần reset.', 'MISSING_FIELDS');

  var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
  var rows = sheetToObjects(sheet);
  var target = rows.filter(function (r) { return String(r.MaNV) === maNV; })[0];
  if (!target) return errorResponse('Không tìm thấy nhân viên.', 'NOT_FOUND');

  var newSalt = generateSalt();
  var newHash = hashPassword(maNV, newSalt);

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colHash = headers.indexOf('PasswordHash') + 1;
  var colSalt = headers.indexOf('PasswordSalt') + 1;
  var colMustChange = headers.indexOf('MustChangePassword') + 1;

  sheet.getRange(target.__row, colHash).setValue(newHash);
  sheet.getRange(target.__row, colSalt).setValue(newSalt);
  sheet.getRange(target.__row, colMustChange).setValue(true);

  writeAudit(session.cccd, 'ADMIN_RESET_PASSWORD', maNV, 'Reset về mã nhân viên');

  return successResponse({ message: 'Đã đặt lại mật khẩu về Mã nhân viên cho ' + target.HoTen + '.' });
}

/**
 * POST /adminSyncStatus
 * body: { token }
 */
function actionAdminSyncStatus(params) {
  var session = requireAdmin(params.token);
  if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');

  var sheet = getOrCreateSheet(CONFIG.SHEET_SYNC_STATUS);
  var rows = sheetToObjects(sheet);
  rows.sort(function (a, b) {
    return new Date(b.LastSyncAt) - new Date(a.LastSyncAt);
  });
  return successResponse({ syncStatus: rows });
}

/**
 * POST /adminAuditLog
 * body: { token, limit }
 */
function actionAdminAuditLog(params) {
  var session = requireAdmin(params.token);
  if (!session) return errorResponse('Bạn không có quyền truy cập.', 'FORBIDDEN');

  var limit = Number(params.limit) || 200;
  var sheet = getOrCreateSheet(CONFIG.SHEET_AUDIT);
  var rows = sheetToObjects(sheet);
  rows.reverse();
  if (rows.length > limit) rows = rows.slice(0, limit);
  return successResponse({ logs: rows });
}

/**
 * Che bớt số CCCD khi hiển thị cho admin trong danh sách (chỉ hiện 4 số cuối),
 * hạn chế lộ dữ liệu nhạy cảm khi nhiều admin cùng xem. Đổi lại nếu bạn muốn
 * hiển thị đầy đủ.
 */
function maskCCCD(cccd) {
  var s = normalizeCCCD(cccd);
  if (s.length <= 4) return s;
  return new Array(s.length - 3).join('*') + s.slice(-4);
}
