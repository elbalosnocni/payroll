/** Entry points for the Google Apps Script Web App. */
function doGet(e) {
  try {
    return jsonOutput({ ok: true, service: 'payroll-api', message: 'Payroll API đang hoạt động.', time: formatDateVN(new Date()) });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message || String(err), code: 'CONFIG_ERROR' });
  }
}

function doPost(e) {
  var params = parseRequestParams(e);
  var action = String(params.action || '').trim();
  try {
    switch (action) {
      case 'ping': return successResponse({ pong: true, time: formatDateVN(new Date()) });
      case 'health': return actionHealth();
      case 'setup': return actionSetup(params);
      case 'login': return actionLogin(params);
      case 'changePassword': return actionChangePassword(params);
      case 'getPayroll': return actionGetPayroll(params);
      case 'adminListEmployees': return actionAdminListEmployees(params);
      case 'adminSearchEmployee': return actionAdminSearchEmployee(params);
      case 'adminResetPassword': return actionAdminResetPassword(params);
      case 'adminSyncStatus': return actionAdminSyncStatus(params);
      case 'adminAuditLog': return actionAdminAuditLog(params);
      case 'sync': return actionSync(params);
      default: return errorResponse('Action không hợp lệ: ' + action, 'UNKNOWN_ACTION');
    }
  } catch (err) {
    Logger.log('Error action=%s: %s\n%s', action, err, err.stack || '');
    return errorResponse('Lỗi máy chủ: ' + (err.message || String(err)), 'INTERNAL_ERROR');
  }
}

function parseRequestParams(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && typeof body === 'object') return body;
    } catch (_) {}
  }
  return (e && e.parameter) ? e.parameter : {};
}

function actionHealth() {
  var cfg = checkConfiguration();
  var result = { configuration: cfg, sheets: [] };
  if (cfg.spreadsheetIdConfigured) {
    var ss = getSS();
    result.spreadsheetName = ss.getName();
    result.sheets = ss.getSheets().map(function(s) { return s.getName(); });
  }
  return successResponse(result);
}

function actionSetup(params) {
  // Setup không được mở công khai: yêu cầu sync key.
  if (!CONFIG.SYNC_API_KEY || !safeCompare(String(params.apiKey || ''), CONFIG.SYNC_API_KEY)) {
    return errorResponse('Sai API key.', 'FORBIDDEN');
  }
  setupSheets();
  return successResponse({ message: 'Đã kiểm tra/tạo đủ các sheet.', sheets: [CONFIG.SHEET_EMPLOYEES, CONFIG.SHEET_PAYROLL, CONFIG.SHEET_AUDIT, CONFIG.SHEET_SYNC_STATUS] });
}
