function doGet() {
  return jsonOutput({ok:true, data:{service:'Payroll API', time:formatDateVN(new Date())}});
}

function doPost(e) {
  var params = parseRequestParams(e);
  var action = String(params.action || '').trim();
  try {
    switch(action) {
      case 'login': return actionLogin(params);
      case 'changePassword': return actionChangePassword(params);
      case 'getPayroll': return actionGetPayroll(params);
      case 'adminListEmployees': return actionAdminListEmployees(params);
      case 'adminSearchEmployee': return actionAdminSearchEmployee(params);
      case 'adminResetPassword': return actionAdminResetPassword(params);
      case 'adminSyncStatus': return actionAdminSyncStatus(params);
      case 'adminAuditLog': return actionAdminAuditLog(params);
      case 'sync': return actionSync(params);
      case 'ping': return successResponse({pong:true});
      default: return errorResponse('Action không hợp lệ.','UNKNOWN_ACTION');
    }
  } catch(err) {
    console.error(err && err.stack ? err.stack : err);
    return errorResponse('Đã xảy ra lỗi hệ thống. Vui lòng thử lại sau.','INTERNAL_ERROR');
  }
}

function parseRequestParams(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && typeof body === 'object') return body;
    } catch(ignore) {}
  }
  return (e && e.parameter) ? e.parameter : {};
}
