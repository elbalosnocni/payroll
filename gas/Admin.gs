function adminGetUsers_(
  token,
  search,
  ipAddress
) {
  const admin =
    requireAdminToken_(token);

  const query =
    normalizeString_(search)
      .toLocaleLowerCase('vi-VN');

  const sheet =
    getSheet_(CONFIG.SHEETS.DATA);

  if (sheet.getLastRow() < 2) {
    return createResponse_(
      true,
      [],
      'Không có nhân viên.',
      'OK'
    );
  }

  const headers =
    getHeaders_(sheet);

  const rows =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        headers.length
      )
      .getDisplayValues();

  const result = [];

  rows.forEach(function(row) {
    const data =
      rowsToObjects_(
        headers,
        [row]
      )[0];

    const employeeCode =
      normalizeEmployeeCode_(
        data.EmployeeCode
      );

    const fullName =
      normalizeString_(
        data.FullName
      );

    const cccd =
      normalizeCCCD_(
        data.CCCD
      );

    const matches =
      !query ||
      employeeCode
        .toLocaleLowerCase('vi-VN')
        .indexOf(query) >= 0 ||
      fullName
        .toLocaleLowerCase('vi-VN')
        .indexOf(query) >= 0 ||
      cccd.indexOf(query) >= 0 ||
      normalizeString_(data.Department)
        .toLocaleLowerCase('vi-VN')
        .indexOf(query) >= 0;

    if (matches) {
      result.push({
        employeeCode: employeeCode,
        fullName: fullName,
        cccd: cccd,
        department: data.Department,
        section: data.Section,
        position: data.Position,
        mustChangePassword:
          String(data.MustChangePassword)
            .toLowerCase() === 'true',
        updatedAt: data.UpdatedAt
      });
    }
  });

  auditLog_(
    'ADMIN',
    admin.sub,
    'ADMIN_GET_USERS',
    query,
    true,
    ipAddress,
    'Returned ' + result.length + ' users'
  );

  return createResponse_(
    true,
    result,
    'Lấy danh sách nhân viên thành công.',
    'OK'
  );
}


function adminResetPassword_(
  token,
  employeeCode,
  ipAddress
) {
  const admin =
    requireAdminToken_(token);

  const code =
    normalizeEmployeeCode_(employeeCode);

  if (!code) {
    throw new Error('INVALID_EMPLOYEE_CODE');
  }

  const employee =
    findEmployeeByCode_(code);

  if (!employee) {
    throw new Error('EMPLOYEE_NOT_FOUND');
  }

  const defaultPassword =
    code;

  const passwordRecord =
    createPasswordRecord_(
      defaultPassword
    );

  const sheet =
    getSheet_(CONFIG.SHEETS.DATA);

  const headers =
    employee.headers;

  const hashColumn =
    findColumnIndex_(
      headers,
      'PasswordHash'
    ) + 1;

  const saltColumn =
    findColumnIndex_(
      headers,
      'PasswordSalt'
    ) + 1;

  const mustChangeColumn =
    findColumnIndex_(
      headers,
      'MustChangePassword'
    ) + 1;

  const updatedAtColumn =
    findColumnIndex_(
      headers,
      'UpdatedAt'
    ) + 1;

  const updatedByColumn =
    findColumnIndex_(
      headers,
      'UpdatedBy'
    ) + 1;

  const lock =
    acquireScriptLock_(10000);

  try {
    sheet
      .getRange(
        employee.rowNumber,
        hashColumn
      )
      .setValue(
        passwordRecord.passwordHash
      );

    sheet
      .getRange(
        employee.rowNumber,
        saltColumn
      )
      .setValue(
        passwordRecord.passwordSalt
      );

    sheet
      .getRange(
        employee.rowNumber,
        mustChangeColumn
      )
      .setValue(true);

    sheet
      .getRange(
        employee.rowNumber,
        updatedAtColumn
      )
      .setValue(
        formatDateTime_(new Date())
      );

    sheet
      .getRange(
        employee.rowNumber,
        updatedByColumn
      )
      .setValue(
        'ADMIN:' + admin.sub
      );
  } finally {
    lock.releaseLock();
  }

  auditLog_(
    'ADMIN',
    admin.sub,
    'ADMIN_RESET_PASSWORD',
    code,
    true,
    ipAddress,
    'Password reset to employee code'
  );

  return createResponse_(
    true,
    {
      employeeCode: code,
      mustChangePassword: true
    },
    'Đã reset mật khẩu.',
    'OK'
  );
}


function adminGetAuditLogs_(
  token,
  search,
  limit,
  ipAddress
) {
  const admin =
    requireAdminToken_(token);

  const sheet =
    getSheet_(CONFIG.SHEETS.AUDIT);

  if (sheet.getLastRow() < 2) {
    return createResponse_(
      true,
      [],
      'Không có audit log.',
      'OK'
    );
  }

  const headers =
    getHeaders_(sheet);

  const rows =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        headers.length
      )
      .getDisplayValues();

  const query =
    normalizeString_(search)
      .toLocaleLowerCase('vi-VN');

  let max =
    Number(limit || 200);

  if (!isFinite(max)) {
    max = 200;
  }

  max =
    Math.max(
      1,
      Math.min(
        500,
        Math.floor(max)
      )
    );

  const result = [];

  for (
    let index = rows.length - 1;
    index >= 0 &&
    result.length < max;
    index--
  ) {
    const data =
      rowsToObjects_(
        headers,
        [rows[index]]
      )[0];

    const searchable =
      JSON.stringify(data)
        .toLocaleLowerCase('vi-VN');

    if (
      !query ||
      searchable.indexOf(query) >= 0
    ) {
      result.push({
        timestamp: data.Timestamp,
        actorType: data.ActorType,
        actorId: data.ActorId,
        action: data.Action,
        target: data.Target,
        success: data.Success,
        details: data.Details
      });
    }
  }

  auditLog_(
    'ADMIN',
    admin.sub,
    'ADMIN_GET_AUDIT',
    query,
    true,
    ipAddress,
    'Returned ' + result.length + ' audit rows'
  );

  return createResponse_(
    true,
    result,
    'Lấy audit log thành công.',
    'OK'
  );
}


function createInitialAdmin_(
  username,
  password
) {
  const user =
    assertString_(
      username,
      'username',
      3,
      100
    );

  const pwd =
    validatePassword_(password);

  const existing =
    findAdminByUsername_(user);

  if (existing) {
    throw new Error('ADMIN_ALREADY_EXISTS');
  }

  const record =
    createPasswordRecord_(pwd);

  const sheet =
    getSheet_(CONFIG.SHEETS.ADMIN);

  const headers =
    getHeaders_(sheet);

  const adminId =
    randomToken_(24);

  const row =
    objectToRow_(
      headers,
      {
        AdminId: adminId,
        Username: user,
        PasswordHash:
          record.passwordHash,
        PasswordSalt:
          record.passwordSalt,
        Active: true,
        CreatedAt:
          formatDateTime_(new Date()),
        UpdatedAt:
          formatDateTime_(new Date())
      }
    );

  sheet
    .appendRow(row);

  auditLog_(
    'SYSTEM',
    'SYSTEM',
    'CREATE_ADMIN',
    user,
    true,
    '',
    'Initial admin created'
  );

  return {
    success: true,
    adminId: adminId,
    username: user
  };
}
