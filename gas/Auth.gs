function findEmployeeByCCCD_(cccd) {
  const normalizedCCCD = normalizeCCCD_(cccd);

  if (!isValidCCCD_(normalizedCCCD)) {
    return null;
  }

  const sheet = getSheet_(CONFIG.SHEETS.DATA);
  const headers = getHeaders_(sheet);

  if (!headers.length || sheet.getLastRow() < 2) {
    return null;
  }

  const cccdIndex = findColumnIndex_(headers, 'CCCD');

  const values = sheet
    .getRange(
      2,
      cccdIndex + 1,
      sheet.getLastRow() - 1,
      1
    )
    .getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    const sheetCCCD = normalizeCCCD_(
      values[index][0]
    );

    if (timingSafeEqual_(
      sheetCCCD,
      normalizedCCCD
    )) {
      const rowNumber = index + 2;

      const row = sheet
        .getRange(
          rowNumber,
          1,
          1,
          headers.length
        )
        .getValues()[0];

      return {
        rowNumber: rowNumber,
        headers: headers,
        data: rowsToObjects_(
          headers,
          [row]
        )[0]
      };
    }
  }

  return null;
}


function findEmployeeByCode_(employeeCode) {
  const normalizedCode =
    normalizeEmployeeCode_(employeeCode);

  if (!normalizedCode) {
    return null;
  }

  const sheet = getSheet_(CONFIG.SHEETS.DATA);
  const headers = getHeaders_(sheet);

  if (!headers.length || sheet.getLastRow() < 2) {
    return null;
  }

  const codeIndex =
    findColumnIndex_(headers, 'EmployeeCode');

  const values = sheet
    .getRange(
      2,
      codeIndex + 1,
      sheet.getLastRow() - 1,
      1
    )
    .getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    if (
      normalizeEmployeeCode_(values[index][0]) ===
      normalizedCode
    ) {
      const rowNumber = index + 2;

      const row = sheet
        .getRange(
          rowNumber,
          1,
          1,
          headers.length
        )
        .getValues()[0];

      return {
        rowNumber: rowNumber,
        headers: headers,
        data: rowsToObjects_(
          headers,
          [row]
        )[0]
      };
    }
  }

  return null;
}


function loginUser_(cccd, password, ipAddress) {
  try {
    const normalizedCCCD = normalizeCCCD_(cccd);
    const suppliedPassword = String(password || '');

    if (!isValidCCCD_(normalizedCCCD)) {
      return createResponse_(
        false,
        null,
        'Thông tin đăng nhập không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    if (
      suppliedPassword.length < 1 ||
      suppliedPassword.length > CONFIG.PASSWORD.MAX_LENGTH
    ) {
      return createResponse_(
        false,
        null,
        'Thông tin đăng nhập không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    const employee = findEmployeeByCCCD_(
      normalizedCCCD
    );

    if (!employee) {
      auditLog_(
        'USER',
        normalizedCCCD,
        'LOGIN',
        normalizedCCCD,
        false,
        ipAddress,
        'Employee not found'
      );

      return createResponse_(
        false,
        null,
        'Thông tin đăng nhập không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    const data = employee.data;

    const employeeCode =
      normalizeEmployeeCode_(data.EmployeeCode);

    const salt = normalizeString_(
      data.PasswordSalt
    );

    const passwordHash = normalizeString_(
      data.PasswordHash
    );

    const mustChangePassword =
      String(data.MustChangePassword)
        .toLowerCase() === 'true';

    let valid = false;

    if (salt && passwordHash) {
      const calculatedHash =
        hashPassword_(
          suppliedPassword,
          salt
        );

      valid = timingSafeEqual_(
        calculatedHash,
        passwordHash
      );
    } else {
      valid = timingSafeEqual_(
        suppliedPassword,
        employeeCode
      );
    }

    if (!valid) {
      auditLog_(
        'USER',
        employeeCode,
        'LOGIN',
        normalizedCCCD,
        false,
        ipAddress,
        'Invalid password'
      );

      return createResponse_(
        false,
        null,
        'Thông tin đăng nhập không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    const token = createUserToken_(
      employeeCode,
      mustChangePassword
    );

    auditLog_(
      'USER',
      employeeCode,
      'LOGIN',
      normalizedCCCD,
      true,
      ipAddress,
      'Successful login'
    );

    return createResponse_(
      true,
      {
        token: token,
        employeeCode: employeeCode,
        fullName: data.FullName,
        mustChangePassword: mustChangePassword
      },
      'Đăng nhập thành công.',
      'OK'
    );
  } catch (error) {
    console.error(error);

    return createResponse_(
      false,
      null,
      'Không thể xử lý đăng nhập.',
      'SERVER_ERROR'
    );
  }
}


function changeUserPassword_(
  token,
  currentPassword,
  newPassword,
  confirmPassword,
  ipAddress
) {
  const payload = requireUserToken_(token);

  const employeeCode =
    normalizeEmployeeCode_(payload.sub);

  const current =
    validatePassword_(currentPassword);

  const next =
    validatePassword_(newPassword);

  const confirm =
    validatePassword_(confirmPassword);

  if (next !== confirm) {
    throw new Error('PASSWORD_CONFIRM_MISMATCH');
  }

  if (current === next) {
    throw new Error('PASSWORD_MUST_BE_DIFFERENT');
  }

  const employee =
    findEmployeeByCode_(employeeCode);

  if (!employee) {
    throw new Error('EMPLOYEE_NOT_FOUND');
  }

  const data = employee.data;

  const currentHash =
    hashPassword_(
      current,
      String(data.PasswordSalt || '')
    );

  if (!timingSafeEqual_(
    currentHash,
    String(data.PasswordHash || '')
  )) {
    auditLog_(
      'USER',
      employeeCode,
      'CHANGE_PASSWORD',
      employeeCode,
      false,
      ipAddress,
      'Current password incorrect'
    );

    throw new Error('CURRENT_PASSWORD_INVALID');
  }

  const passwordRecord =
    createPasswordRecord_(next);

  const sheet =
    getSheet_(CONFIG.SHEETS.DATA);

  const headers =
    employee.headers;

  const hashColumn =
    findColumnIndex_(headers, 'PasswordHash') + 1;

  const saltColumn =
    findColumnIndex_(headers, 'PasswordSalt') + 1;

  const mustChangeColumn =
    findColumnIndex_(headers, 'MustChangePassword') + 1;

  const updatedAtColumn =
    findColumnIndex_(headers, 'UpdatedAt') + 1;

  const updatedByColumn =
    findColumnIndex_(headers, 'UpdatedBy') + 1;

  const lock = acquireScriptLock_(10000);

  try {
    sheet
      .getRange(employee.rowNumber, hashColumn)
      .setValue(passwordRecord.passwordHash);

    sheet
      .getRange(employee.rowNumber, saltColumn)
      .setValue(passwordRecord.passwordSalt);

    sheet
      .getRange(employee.rowNumber, mustChangeColumn)
      .setValue(false);

    sheet
      .getRange(employee.rowNumber, updatedAtColumn)
      .setValue(formatDateTime_(new Date()));

    sheet
      .getRange(employee.rowNumber, updatedByColumn)
      .setValue(employeeCode);

    auditLog_(
      'USER',
      employeeCode,
      'CHANGE_PASSWORD',
      employeeCode,
      true,
      ipAddress,
      'Password changed'
    );
  } finally {
    lock.releaseLock();
  }

  const newToken =
    createUserToken_(
      employeeCode,
      false
    );

  return createResponse_(
    true,
    {
      token: newToken,
      employeeCode: employeeCode,
      mustChangePassword: false
    },
    'Đổi mật khẩu thành công.',
    'OK'
  );
}


function findAdminByUsername_(username) {
  const normalized =
    normalizeString_(username)
      .toLowerCase();

  const sheet =
    getSheet_(CONFIG.SHEETS.ADMIN);

  if (sheet.getLastRow() < 2) {
    return null;
  }

  const headers =
    getHeaders_(sheet);

  const usernameIndex =
    findColumnIndex_(headers, 'Username');

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        headers.length
      )
      .getValues();

  for (let index = 0; index < values.length; index++) {
    const rowObject =
      rowsToObjects_(
        headers,
        [values[index]]
      )[0];

    if (
      normalizeString_(rowObject.Username)
        .toLowerCase() === normalized &&
      String(rowObject.Active).toLowerCase() === 'true'
    ) {
      return {
        rowNumber: index + 2,
        headers: headers,
        data: rowObject
      };
    }
  }

  return null;
}


function loginAdmin_(
  username,
  password,
  ipAddress
) {
  try {
    const admin =
      findAdminByUsername_(username);

    if (!admin) {
      auditLog_(
        'ADMIN',
        username,
        'ADMIN_LOGIN',
        username,
        false,
        ipAddress,
        'Admin not found'
      );

      return createResponse_(
        false,
        null,
        'Thông tin quản trị không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    const data = admin.data;

    const calculated =
      hashPassword_(
        String(password || ''),
        String(data.PasswordSalt || '')
      );

    if (!timingSafeEqual_(
      calculated,
      String(data.PasswordHash || '')
    )) {
      auditLog_(
        'ADMIN',
        username,
        'ADMIN_LOGIN',
        username,
        false,
        ipAddress,
        'Invalid admin password'
      );

      return createResponse_(
        false,
        null,
        'Thông tin quản trị không hợp lệ.',
        'INVALID_CREDENTIALS'
      );
    }

    const token =
      createAdminToken_(
        data.AdminId,
        data.Username
      );

    auditLog_(
      'ADMIN',
      data.AdminId,
      'ADMIN_LOGIN',
      data.Username,
      true,
      ipAddress,
      'Successful admin login'
    );

    return createResponse_(
      true,
      {
        token: token,
        username: data.Username
      },
      'Đăng nhập quản trị thành công.',
      'OK'
    );
  } catch (error) {
    console.error(error);

    return createResponse_(
      false,
      null,
      'Không thể xử lý đăng nhập quản trị.',
      'SERVER_ERROR'
    );
  }
}
