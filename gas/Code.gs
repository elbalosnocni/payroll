function jsonOutput_(object) {
  return ContentService
    .createTextOutput(
      JSON.stringify(object)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


function jsonpOutput_(
  callback,
  object
) {
  const safeCallback =
    /^[A-Za-z_$][A-Za-z0-9_$]*(\.[A-Za-z_$][A-Za-z0-9_$]*)*$/
      .test(String(callback || ''))
      ? String(callback)
      : 'callback';

  return ContentService
    .createTextOutput(
      safeCallback +
      '(' +
      JSON.stringify(object) +
      ')'
    )
    .setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
}


function getParameter_(e, name) {
  if (!e) {
    return '';
  }

  if (e.parameter && e.parameter[name] !== undefined) {
    return e.parameter[name];
  }

  return '';
}


function parseRequestBody_(e) {
  if (!e || !e.postData) {
    return {};
  }

  const contents =
    e.postData.contents || '';

  if (
    String(contents).length >
    CONFIG.API.MAX_BODY_BYTES
  ) {
    throw new Error('REQUEST_TOO_LARGE');
  }

  return safeJsonParse_(contents);
}


function routeRequest_(method, e) {
  const action =
    String(
      getParameter_(e, 'action') ||
      (
        e &&
        e.parameter &&
        e.parameter.route
      ) ||
      ''
    ).trim();

  if (!action) {
    return createResponse_(
      false,
      null,
      'Thiếu action.',
      'MISSING_ACTION'
    );
  }

  const ipAddress =
    getRequestClientIp_(e);

  try {
    if (
      method === 'GET' &&
      action === 'login'
    ) {
      const cccd =
        getParameter_(e, 'cccd');

      const password =
        getParameter_(e, 'password');

      const result =
        loginUser_(
          cccd,
          password,
          ipAddress
        );

      const callback =
        getParameter_(e, 'callback');

      return callback
        ? jsonpOutput_(callback, result)
        : jsonOutput_(result);
    }

    if (
      method === 'GET' &&
      action === 'getPayroll'
    ) {
      const token =
        getParameter_(e, 'token');

      const result =
        getPayrollForCurrentUser_(
          token,
          ipAddress
        );

      const callback =
        getParameter_(e, 'callback');

      return callback
        ? jsonpOutput_(callback, result)
        : jsonOutput_(result);
    }

    if (
      method === 'GET' &&
      action === 'adminGetUsers'
    ) {
      const token =
        getParameter_(e, 'token');

      const search =
        getParameter_(e, 'search');

      const result =
        adminGetUsers_(
          token,
          search,
          ipAddress
        );

      const callback =
        getParameter_(e, 'callback');

      return callback
        ? jsonpOutput_(callback, result)
        : jsonOutput_(result);
    }

    if (
      method === 'GET' &&
      action === 'adminGetAudit'
    ) {
      const token =
        getParameter_(e, 'token');

      const search =
        getParameter_(e, 'search');

      const limit =
        getParameter_(e, 'limit');

      const result =
        adminGetAuditLogs_(
          token,
          search,
          limit,
          ipAddress
        );

      const callback =
        getParameter_(e, 'callback');

      return callback
        ? jsonpOutput_(callback, result)
        : jsonOutput_(result);
    }

    const body =
      method === 'POST'
        ? parseRequestBody_(e)
        : {};

    if (
      method === 'POST' &&
      action === 'changePassword'
    ) {
      const result =
        changeUserPassword_(
          body.token,
          body.currentPassword,
          body.newPassword,
          body.confirmPassword,
          ipAddress
        );

      return jsonOutput_(result);
    }

    if (
      method === 'POST' &&
      action === 'adminLogin'
    ) {
      const result =
        loginAdmin_(
          body.username,
          body.password,
          ipAddress
        );

      return jsonOutput_(result);
    }

    if (
      method === 'POST' &&
      action === 'adminResetPwd'
    ) {
      const result =
        adminResetPassword_(
          body.token,
          body.employeeCode,
          ipAddress
        );

      return jsonOutput_(result);
    }

    if (
      method === 'POST' &&
      action === 'syncVBA'
    ) {
      const result =
        syncPayrollData_(
          body.syncKey,
          body.records,
          body.sourceFile,
          body.sourceMonth,
          ipAddress
        );

      return jsonOutput_(result);
    }

    return jsonOutput_(
      createResponse_(
        false,
        null,
        'Action không được hỗ trợ.',
        'UNKNOWN_ACTION'
      )
    );
  } catch (error) {
    console.error(error);

    const message =
      mapErrorMessage_(error);

    const result =
      createResponse_(
        false,
        null,
        message.message,
        message.code
      );

    return jsonOutput_(result);
  }
}


function mapErrorMessage_(error) {
  const code =
    String(
      error &&
      error.message
        ? error.message
        : 'SERVER_ERROR'
    );

  const messages = {
    UNAUTHORIZED:
      'Phiên đăng nhập không hợp lệ hoặc đã hết hạn.',

    ADMIN_UNAUTHORIZED:
      'Phiên quản trị không hợp lệ hoặc đã hết hạn.',

    INVALID_CREDENTIALS:
      'Thông tin đăng nhập không hợp lệ.',

    CURRENT_PASSWORD_INVALID:
      'Mật khẩu hiện tại không đúng.',

    PASSWORD_CONFIRM_MISMATCH:
      'Mật khẩu xác nhận không khớp.',

    PASSWORD_MUST_BE_DIFFERENT:
      'Mật khẩu mới phải khác mật khẩu hiện tại.',

    INVALID_PASSWORD_LENGTH:
      'Mật khẩu phải có độ dài từ 8 đến 128 ký tự.',

    EMPLOYEE_NOT_FOUND:
      'Không tìm thấy nhân viên.',

    INVALID_EMPLOYEE_CODE:
      'Mã nhân viên không hợp lệ.',

    INVALID_SYNC_KEY:
      'Khóa đồng bộ không hợp lệ.',

    INVALID_SYNC_RECORDS:
      'Dữ liệu đồng bộ không hợp lệ.',

    INVALID_SYNC_RECORD_COUNT:
      'Số lượng bản ghi đồng bộ không hợp lệ.',

    SERVER_BUSY:
      'Hệ thống đang xử lý yêu cầu khác. Vui lòng thử lại.',

    REQUEST_TOO_LARGE:
      'Dữ liệu gửi lên vượt quá giới hạn.',

    ADMIN_ALREADY_EXISTS:
      'Tài khoản quản trị đã tồn tại.'
  };

  return {
    code: code,
    message:
      messages[code] ||
      'Có lỗi xảy ra trong quá trình xử lý.'
  };
}


function doGet(e) {
  try {
    return routeRequest_(
      'GET',
      e
    );
  } catch (error) {
    return jsonOutput_(
      createResponse_(
        false,
        null,
        'Internal server error.',
        'SERVER_ERROR'
      )
    );
  }
}


function doPost(e) {
  try {
    return routeRequest_(
      'POST',
      e
    );
  } catch (error) {
    return jsonOutput_(
      createResponse_(
        false,
        null,
        'Internal server error.',
        'SERVER_ERROR'
      )
    );
  }
}
