/**
 * Code.gs
 * -----------------------------------------------------------------------
 * Điểm vào chính của Apps Script Web App. Nhận request POST (JSON body hoặc
 * form-urlencoded) chứa "action" để quyết định gọi hàm nào.
 *
 * Cách deploy: xem docs/GAS_DEPLOY.md
 * -----------------------------------------------------------------------
 */

function doGet(e) {
  return successResponse({
    message: 'Payroll API đang hoạt động.',
    time: formatDateVN(new Date())
  });
}

function doPost(e) {
  var params = parseRequestParams(e);
  var action = params.action;

  try {
    switch (action) {
      case 'login':
        return actionLogin(params);
      case 'changePassword':
        return actionChangePassword(params);
      case 'getPayroll':
        return actionGetPayroll(params);
      case 'adminListEmployees':
        return actionAdminListEmployees(params);
      case 'adminSearchEmployee':
        return actionAdminSearchEmployee(params);
      case 'adminResetPassword':
        return actionAdminResetPassword(params);
      case 'adminSyncStatus':
        return actionAdminSyncStatus(params);
      case 'adminAuditLog':
        return actionAdminAuditLog(params);
      case 'sync':
        return actionSync(params);
      case 'ping':
        return successResponse({ pong: true });
      default:
        return errorResponse('Action không hợp lệ: ' + action, 'UNKNOWN_ACTION');
    }
  } catch (err) {
    // Không lộ chi tiết lỗi hệ thống ra ngoài, chỉ log lại để debug qua
    // Apps Script > Executions.
    Logger.log('Error in action ' + action + ': ' + err + '\n' + err.stack);
    return errorResponse('Đã xảy ra lỗi hệ thống. Vui lòng thử lại sau.', 'INTERNAL_ERROR');
  }
}

/**
 * Đọc tham số từ request: ưu tiên JSON body (e.postData.contents), fallback
 * sang e.parameter (form-urlencoded) nếu cần.
 */
function parseRequestParams(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && typeof body === 'object') return body;
    } catch (err) {
      // không phải JSON, rơi xuống dùng e.parameter
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}
