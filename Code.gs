const CONFIG = {
  SPREADSHEET_ID: '1Y-f3xSzwrV2ycBnzThlrtgcOcHD2Z4iqLKqqTyVFejM',

  SESSION_SECONDS: 1800,

  SYNC_API_KEY: 'TRUONGCONGVU_SALARYSLIP_1988_Elbalosnocni',

  SETUP_KEY: '1235678',

  DEFAULT_ADMIN_USERNAME: 'admin',

  DEFAULT_ADMIN_PASSWORD: 'Admin@123456',

  TIMEZONE: 'Asia/Ho_Chi_Minh',

  SHEETS: {
    EMPLOYEES: 'Employees',
    PAYROLL: 'Payroll',
    ADMINS: 'Admins',
    SESSIONS: 'Sessions',
    AUDIT: 'AuditLogs',
    SYNC: 'SyncStatus'
  }
};

const EMPLOYEE_HEADERS = [
  'EmployeeCode',
  'FullName',
  'CitizenID',
  'Department',
  'Section',
  'Position',
  'PasswordHash',
  'PasswordSalt',
  'MustChangePassword',
  'Active',
  'UpdatedAt'
];

const PAYROLL_HEADERS = [
  'EmployeeCode',
  'FullName',
  'CitizenID',
  'Department',
  'Section',
  'Position',
  'SalaryMonth',

  'BasicSalary',
  'WorkingDays',
  'HolidayDays',
  'PaidLeaveDays',
  'UnpaidLeaveDays',

  'OvertimeHours',
  'HolidayWorkHours',
  'HolidayOvertimeHours',

  'MinimumRegionalLeaveDays',
  'NightHolidayOvertimeHours',
  'NightOvertimeHours',

  'HolidayWorkDayHours',
  'HolidayOvertimeDayHours',
  'NightShiftDays',

  'OtherMoney',
  'DisciplinaryMoney',
  'Loyalty2Years',
  'Loyalty5Years',
  'Loyalty10Years',
  'HousingAllowance',
  'TransportationAllowance',
  'AttendanceBonus',
  'SeveranceAndUnusedLeave',

  'MonthlySalary',
  'OvertimeSalary',
  'CommissionAndOverTargetBonus',
  'OtherIncome',

  'OtherDeductions',
  'AdvancePayment',

  'SocialInsurance',
  'HealthInsurance',
  'UnemploymentInsurance',
  'PersonalIncomeTax',

  'GrossIncome',
  'NetSalary',

  'UpdatedAt'
];

const ADMIN_HEADERS = [
  'Username',
  'PasswordHash',
  'PasswordSalt',
  'Active',
  'UpdatedAt'
];

const SESSION_HEADERS = [
  'Token',
  'Username',
  'Role',
  'EmployeeCode',
  'CreatedAt',
  'ExpiresAt',
  'Active'
];

const AUDIT_HEADERS = [
  'Timestamp',
  'Username',
  'Role',
  'Action',
  'Target',
  'IP',
  'Details'
];

const SYNC_HEADERS = [
  'Timestamp',
  'Status',
  'Month',
  'SnackFile',
  'FlexibleFile',
  'EmployeeCount',
  'PayrollCount',
  'Message'
];

function doGet(e) {
  return jsonResponse_({
    success: true,
    service: 'Internal Payroll API',
    version: '1.0.0',
    timestamp: nowString_()
  });
}

function doPost(e) {
  try {
    const body = parseRequest_(e);

    const action = String(body.action || '').trim();

    if (!action) {
      return jsonResponse_({
        success: false,
        error: 'ACTION_REQUIRED'
      });
    }

    switch (action) {
      case 'login':
        return jsonResponse_(login_(body));

      case 'changePassword':
        return jsonResponse_(changePassword_(body));

      case 'getPayroll':
        return jsonResponse_(getPayroll_(body));

      case 'logout':
        return jsonResponse_(logout_(body));

      case 'adminLogin':
        return jsonResponse_(adminLogin_(body));

      case 'adminEmployees':
        return jsonResponse_(adminEmployees_(body));

      case 'adminSearchEmployees':
        return jsonResponse_(adminSearchEmployees_(body));

      case 'adminResetPassword':
        return jsonResponse_(adminResetPassword_(body));

      case 'adminAuditLogs':
        return jsonResponse_(adminAuditLogs_(body));

      case 'adminSyncStatus':
        return jsonResponse_(adminSyncStatus_(body));

      case 'syncPayroll':
        return jsonResponse_(syncPayroll_(body));

      case 'setup':
        return jsonResponse_(setup_(body));

      default:
        return jsonResponse_({
          success: false,
          error: 'UNKNOWN_ACTION'
        });
    }
  } catch (error) {
    console.error(error);

    return jsonResponse_({
      success: false,
      error: 'SERVER_ERROR',
      message: String(error.message || error)
    });
  }
}

function setup_(body) {
  if (String(body.setupKey || '') !== CONFIG.SETUP_KEY) {
    return {
      success: false,
      error: 'INVALID_SETUP_KEY'
    };
  }

  ensureSystem_();

  const sheet = getSheet_(CONFIG.SHEETS.ADMINS);

  const values = sheet.getDataRange().getValues();

  let exists = false;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === CONFIG.DEFAULT_ADMIN_USERNAME) {
      exists = true;
      break;
    }
  }

  if (!exists) {
    const passwordData = createPasswordHash_(
      CONFIG.DEFAULT_ADMIN_PASSWORD
    );

    sheet.appendRow([
      CONFIG.DEFAULT_ADMIN_USERNAME,
      passwordData.hash,
      passwordData.salt,
      true,
      nowString_()
    ]);
  }

  writeAudit_(
    CONFIG.DEFAULT_ADMIN_USERNAME,
    'ADMIN',
    'SYSTEM_SETUP',
    '',
    '',
    'System initialized'
  );

  return {
    success: true,
    message: 'SYSTEM_READY',
    adminUsername: CONFIG.DEFAULT_ADMIN_USERNAME
  };
}

function ensureSystem_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  ensureSheet_(ss, CONFIG.SHEETS.EMPLOYEES, EMPLOYEE_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.PAYROLL, PAYROLL_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.ADMINS, ADMIN_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, SESSION_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.AUDIT, AUDIT_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.SYNC, SYNC_HEADERS);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const current = sheet
      .getRange(1, 1, 1, headers.length)
      .getValues()[0];

    let changed = false;

    for (let i = 0; i < headers.length; i++) {
      if (String(current[i] || '') !== headers[i]) {
        changed = true;
        break;
      }
    }

    if (changed) {
      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers]);
    }
  }

  return sheet;
}

function login_(body) {
  ensureSystem_();

  const citizenID = normalizeCitizenID_(body.citizenID);
  const password = String(body.password || '');

  if (!citizenID || !password) {
    return {
      success: false,
      error: 'INVALID_LOGIN'
    };
  }

  const employee = findEmployeeByCitizenID_(citizenID);

  if (!employee || !employee.active) {
    writeAudit_(
      citizenID,
      'EMPLOYEE',
      'LOGIN_FAILED',
      citizenID,
      '',
      'Employee not found or inactive'
    );

    return {
      success: false,
      error: 'INVALID_LOGIN'
    };
  }

  const valid = verifyPassword_(
    password,
    employee.passwordHash,
    employee.passwordSalt
  );

  if (!valid) {
    writeAudit_(
      employee.employeeCode,
      'EMPLOYEE',
      'LOGIN_FAILED',
      employee.employeeCode,
      '',
      'Invalid password'
    );

    return {
      success: false,
      error: 'INVALID_LOGIN'
    };
  }

  const token = createSession_(
    employee.employeeCode,
    'EMPLOYEE'
  );

  writeAudit_(
    employee.employeeCode,
    'EMPLOYEE',
    'LOGIN',
    employee.employeeCode,
    '',
    'Employee login'
  );

  const response = {
    success: true,
    token: token,
    mustChangePassword: employee.mustChangePassword,
    employee: {
      employeeCode: employee.employeeCode,
      fullName: employee.fullName,
      citizenID: employee.citizenID,
      department: employee.department,
      section: employee.section,
      position: employee.position
    }
  };

  if (!employee.mustChangePassword) {
    response.payroll = getLatestPayrollForEmployee_(
      employee.employeeCode
    );
  }

  return response;
}

function changePassword_(body) {
  const session = requireEmployeeSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'SESSION_EXPIRED'
    };
  }

  const currentPassword = String(
    body.currentPassword || ''
  );

  const newPassword = String(
    body.newPassword || ''
  );

  const confirmPassword = String(
    body.confirmPassword || ''
  );

  if (!currentPassword || !newPassword || !confirmPassword) {
    return {
      success: false,
      error: 'PASSWORD_REQUIRED'
    };
  }

  if (newPassword !== confirmPassword) {
    return {
      success: false,
      error: 'PASSWORD_CONFIRM_MISMATCH'
    };
  }

  if (newPassword.length < 8) {
    return {
      success: false,
      error: 'PASSWORD_TOO_SHORT'
    };
  }

  if (currentPassword === newPassword) {
    return {
      success: false,
      error: 'PASSWORD_MUST_BE_DIFFERENT'
    };
  }

  const employee = findEmployeeByEmployeeCode_(
    session.employeeCode
  );

  if (!employee) {
    return {
      success: false,
      error: 'EMPLOYEE_NOT_FOUND'
    };
  }

  if (
    !verifyPassword_(
      currentPassword,
      employee.passwordHash,
      employee.passwordSalt
    )
  ) {
    return {
      success: false,
      error: 'CURRENT_PASSWORD_INVALID'
    };
  }

  const passwordData = createPasswordHash_(newPassword);

  const sheet = getSheet_(CONFIG.SHEETS.EMPLOYEES);

  sheet
    .getRange(employee.row, 7, 1, 4)
    .setValues([[
      passwordData.hash,
      passwordData.salt,
      false,
      true
    ]]);

  sheet
    .getRange(employee.row, 11)
    .setValue(nowString_());

  writeAudit_(
    employee.employeeCode,
    'EMPLOYEE',
    'CHANGE_PASSWORD',
    employee.employeeCode,
    '',
    'Employee changed password'
  );

  return {
    success: true,
    message: 'PASSWORD_CHANGED'
  };
}

function getPayroll_(body) {
  const session = requireEmployeeSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'SESSION_EXPIRED'
    };
  }

  const employee = findEmployeeByEmployeeCode_(
    session.employeeCode
  );

  if (!employee) {
    return {
      success: false,
      error: 'EMPLOYEE_NOT_FOUND'
    };
  }

  if (employee.mustChangePassword) {
    return {
      success: false,
      error: 'PASSWORD_CHANGE_REQUIRED'
    };
  }

  const payroll = getLatestPayrollForEmployee_(
    employee.employeeCode
  );

  return {
    success: true,
    payroll: payroll
  };
}

function logout_(body) {
  const token = String(body.token || '');

  if (!token) {
    return {
      success: true
    };
  }

  const sheet = getSheet_(CONFIG.SHEETS.SESSIONS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === token) {
      sheet.getRange(i + 1, 7).setValue(false);
      break;
    }
  }

  return {
    success: true
  };
}

function adminLogin_(body) {
  ensureSystem_();

  const username = String(body.username || '').trim();
  const password = String(body.password || '');

  const admin = findAdmin_(username);

  if (!admin || !admin.active) {
    writeAudit_(
      username,
      'ADMIN',
      'ADMIN_LOGIN_FAILED',
      username,
      '',
      'Admin not found'
    );

    return {
      success: false,
      error: 'INVALID_LOGIN'
    };
  }

  if (
    !verifyPassword_(
      password,
      admin.passwordHash,
      admin.passwordSalt
    )
  ) {
    writeAudit_(
      username,
      'ADMIN',
      'ADMIN_LOGIN_FAILED',
      username,
      '',
      'Invalid password'
    );

    return {
      success: false,
      error: 'INVALID_LOGIN'
    };
  }

  const token = createSession_(
    username,
    'ADMIN'
  );

  writeAudit_(
    username,
    'ADMIN',
    'ADMIN_LOGIN',
    username,
    '',
    'Admin login'
  );

  return {
    success: true,
    token: token,
    role: 'ADMIN',
    username: username
  };
}

function adminEmployees_(body) {
  const session = requireAdminSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'UNAUTHORIZED'
    };
  }

  const employees = readEmployees_();

  writeAudit_(
    session.username,
    'ADMIN',
    'VIEW_EMPLOYEES',
    '',
    '',
    'Admin viewed employee list'
  );

  return {
    success: true,
    employees: employees
  };
}

function adminSearchEmployees_(body) {
  const session = requireAdminSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'UNAUTHORIZED'
    };
  }

  const keyword = String(
    body.keyword || ''
  ).trim().toLowerCase();

  const employees = readEmployees_();

  const filtered = employees.filter(function(employee) {
    return (
      String(employee.employeeCode).toLowerCase().includes(keyword) ||
      String(employee.fullName).toLowerCase().includes(keyword) ||
      String(employee.citizenID).toLowerCase().includes(keyword) ||
      String(employee.department).toLowerCase().includes(keyword) ||
      String(employee.section).toLowerCase().includes(keyword)
    );
  });

  return {
    success: true,
    employees: filtered
  };
}

function adminResetPassword_(body) {
  const session = requireAdminSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'UNAUTHORIZED'
    };
  }

  const employeeCode = String(
    body.employeeCode || ''
  ).trim();

  if (!employeeCode) {
    return {
      success: false,
      error: 'EMPLOYEE_CODE_REQUIRED'
    };
  }

  const employee = findEmployeeByEmployeeCode_(
    employeeCode
  );

  if (!employee) {
    return {
      success: false,
      error: 'EMPLOYEE_NOT_FOUND'
    };
  }

  const defaultPassword = employee.employeeCode;

  const passwordData = createPasswordHash_(
    defaultPassword
  );

  const sheet = getSheet_(CONFIG.SHEETS.EMPLOYEES);

  sheet
    .getRange(employee.row, 7, 1, 4)
    .setValues([[
      passwordData.hash,
      passwordData.salt,
      true,
      true
    ]]);

  sheet
    .getRange(employee.row, 11)
    .setValue(nowString_());

  writeAudit_(
    session.username,
    'ADMIN',
    'RESET_PASSWORD',
    employeeCode,
    '',
    'Password reset to employee code'
  );

  return {
    success: true,
    message: 'PASSWORD_RESET',
    employeeCode: employeeCode
  };
}

function adminAuditLogs_(body) {
  const session = requireAdminSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'UNAUTHORIZED'
    };
  }

  const sheet = getSheet_(CONFIG.SHEETS.AUDIT);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return {
      success: true,
      logs: []
    };
  }

  const startRow = Math.max(
    2,
    lastRow - 499
  );

  const numberOfRows =
    lastRow - startRow + 1;

  const values = sheet
    .getRange(
      startRow,
      1,
      numberOfRows,
      AUDIT_HEADERS.length
    )
    .getValues();

  const logs = values.reverse().map(function(row) {
    return {
      timestamp: row[0],
      username: row[1],
      role: row[2],
      action: row[3],
      target: row[4],
      ip: row[5],
      details: row[6]
    };
  });

  return {
    success: true,
    logs: logs
  };
}

function adminSyncStatus_(body) {
  const session = requireAdminSession_(body.token);

  if (!session) {
    return {
      success: false,
      error: 'UNAUTHORIZED'
    };
  }

  const sheet = getSheet_(CONFIG.SHEETS.SYNC);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return {
      success: true,
      status: null
    };
  }

  const row = sheet
    .getRange(
      lastRow,
      1,
      1,
      SYNC_HEADERS.length
    )
    .getValues()[0];

  return {
    success: true,
    status: {
      timestamp: row[0],
      status: row[1],
      month: row[2],
      snackFile: row[3],
      flexibleFile: row[4],
      employeeCount: row[5],
      payrollCount: row[6],
      message: row[7]
    }
  };
}

function syncPayroll_(body) {
  if (
    String(body.apiKey || '') !==
    CONFIG.SYNC_API_KEY
  ) {
    return {
      success: false,
      error: 'INVALID_SYNC_KEY'
    };
  }

  ensureSystem_();

  const month = String(
    body.salaryMonth || ''
  ).trim();

  const records = Array.isArray(body.records)
    ? body.records
    : [];

  const source = body.source || {};

  if (!month) {
    return {
      success: false,
      error: 'SALARY_MONTH_REQUIRED'
    };
  }

  if (records.length === 0) {
    return {
      success: false,
      error: 'NO_RECORDS'
    };
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const employeeSheet =
      getSheet_(CONFIG.SHEETS.EMPLOYEES);

    const payrollSheet =
      getSheet_(CONFIG.SHEETS.PAYROLL);

    const existingEmployees =
      readEmployeeIndex_();

    const payrollRows = [];

    let employeeCount = 0;

    records.forEach(function(record) {
      const employeeCode =
        String(record.employeeCode || '').trim();

      const fullName =
        String(record.fullName || '').trim();

      const citizenID =
        normalizeCitizenID_(record.citizenID);

      if (!employeeCode || !fullName) {
        return;
      }

      const employeeData = {
        employeeCode: employeeCode,
        fullName: fullName,
        citizenID: citizenID,
        department: String(record.department || ''),
        section: String(record.section || ''),
        position: String(record.position || '')
      };

      upsertEmployee_(
        employeeSheet,
        existingEmployees,
        employeeData
      );

      employeeCount++;

      payrollRows.push(
        payrollRecordToRow_(
          record,
          month,
          employeeData
        )
      );
    });

    deletePayrollMonth_(month);

    if (payrollRows.length > 0) {
      payrollSheet
        .getRange(
          payrollSheet.getLastRow() + 1,
          1,
          payrollRows.length,
          PAYROLL_HEADERS.length
        )
        .setValues(payrollRows);
    }

    const statusSheet =
      getSheet_(CONFIG.SHEETS.SYNC);

    statusSheet.appendRow([
      nowString_(),
      'SUCCESS',
      month,
      source.snackFile || '',
      source.flexibleFile || '',
      employeeCount,
      payrollRows.length,
      'Payroll synchronized successfully'
    ]);

    writeAudit_(
      'VBA_SYNC',
      'SYSTEM',
      'SYNC_PAYROLL',
      month,
      '',
      'Payroll synchronization completed'
    );

    return {
      success: true,
      month: month,
      employeeCount: employeeCount,
      payrollCount: payrollRows.length
    };
  } catch (error) {
    const statusSheet =
      getSheet_(CONFIG.SHEETS.SYNC);

    statusSheet.appendRow([
      nowString_(),
      'ERROR',
      month,
      source.snackFile || '',
      source.flexibleFile || '',
      0,
      0,
      String(error.message || error)
    ]);

    throw error;
  } finally {
    lock.releaseLock();
  }
}

function payrollRecordToRow_(
  record,
  month,
  employee
) {
  return [
    employee.employeeCode,
    employee.fullName,
    employee.citizenID,
    employee.department,
    employee.section,
    employee.position,
    month,

    numberValue_(record.basicSalary),
    numberValue_(record.workingDays),
    numberValue_(record.holidayDays),
    numberValue_(record.paidLeaveDays),
    numberValue_(record.unpaidLeaveDays),

    numberValue_(record.overtimeHours),
    numberValue_(record.holidayWorkHours),
    numberValue_(record.holidayOvertimeHours),

    numberValue_(record.minimumRegionalLeaveDays),
    numberValue_(record.nightHolidayOvertimeHours),
    numberValue_(record.nightOvertimeHours),

    numberValue_(record.holidayWorkDayHours),
    numberValue_(record.holidayOvertimeDayHours),
    numberValue_(record.nightShiftDays),

    numberValue_(record.otherMoney),
    numberValue_(record.disciplinaryMoney),
    numberValue_(record.loyalty2Years),
    numberValue_(record.loyalty5Years),
    numberValue_(record.loyalty10Years),
    numberValue_(record.housingAllowance),
    numberValue_(record.transportationAllowance),
    numberValue_(record.attendanceBonus),
    numberValue_(record.severanceAndUnusedLeave),

    numberValue_(record.monthlySalary),
    numberValue_(record.overtimeSalary),
    numberValue_(record.commissionAndOverTargetBonus),
    numberValue_(record.otherIncome),

    numberValue_(record.otherDeductions),
    numberValue_(record.advancePayment),

    numberValue_(record.socialInsurance),
    numberValue_(record.healthInsurance),
    numberValue_(record.unemploymentInsurance),
    numberValue_(record.personalIncomeTax),

    numberValue_(record.grossIncome),
    numberValue_(record.netSalary),

    nowString_()
  ];
}

function deletePayrollMonth_(month) {
  const sheet = getSheet_(CONFIG.SHEETS.PAYROLL);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      PAYROLL_HEADERS.length
    )
    .getValues();

  const rowsToDelete = [];

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][6]).trim() === month
    ) {
      rowsToDelete.push(i + 2);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

function upsertEmployee_(
  sheet,
  index,
  employee
) {
  const existing =
    index[employee.employeeCode];

  if (!existing) {
    const passwordData =
      createPasswordHash_(
        employee.employeeCode
      );

    sheet.appendRow([
      employee.employeeCode,
      employee.fullName,
      employee.citizenID,
      employee.department,
      employee.section,
      employee.position,
      passwordData.hash,
      passwordData.salt,
      true,
      true,
      nowString_()
    ]);

    return;
  }

  sheet
    .getRange(existing.row, 1, 1, 6)
    .setValues([[
      employee.employeeCode,
      employee.fullName,
      employee.citizenID,
      employee.department,
      employee.section,
      employee.position
    ]]);

  sheet
    .getRange(existing.row, 11)
    .setValue(nowString_());
}

function readEmployeeIndex_() {
  const sheet =
    getSheet_(CONFIG.SHEETS.EMPLOYEES);

  const lastRow = sheet.getLastRow();

  const result = {};

  if (lastRow <= 1) {
    return result;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      EMPLOYEE_HEADERS.length
    )
    .getValues();

  values.forEach(function(row, index) {
    const employeeCode =
      String(row[0] || '').trim();

    if (!employeeCode) {
      return;
    }

    result[employeeCode] = {
      row: index + 2
    };
  });

  return result;
}

function findEmployeeByCitizenID_(citizenID) {
  const sheet =
    getSheet_(CONFIG.SHEETS.EMPLOYEES);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      EMPLOYEE_HEADERS.length
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const current =
      normalizeCitizenID_(values[i][2]);

    if (current === citizenID) {
      return employeeFromRow_(
        values[i],
        i + 2
      );
    }
  }

  return null;
}

function findEmployeeByEmployeeCode_(code) {
  const sheet =
    getSheet_(CONFIG.SHEETS.EMPLOYEES);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      EMPLOYEE_HEADERS.length
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0]).trim() ===
      String(code).trim()
    ) {
      return employeeFromRow_(
        values[i],
        i + 2
      );
    }
  }

  return null;
}

function employeeFromRow_(row, rowNumber) {
  return {
    row: rowNumber,
    employeeCode: String(row[0] || ''),
    fullName: String(row[1] || ''),
    citizenID: normalizeCitizenID_(row[2]),
    department: String(row[3] || ''),
    section: String(row[4] || ''),
    position: String(row[5] || ''),
    passwordHash: String(row[6] || ''),
    passwordSalt: String(row[7] || ''),
    mustChangePassword:
      row[8] === true ||
      String(row[8]).toLowerCase() === 'true',
    active:
      row[9] === true ||
      String(row[9]).toLowerCase() === 'true',
    updatedAt: String(row[10] || '')
  };
}

function readEmployees_() {
  const sheet =
    getSheet_(CONFIG.SHEETS.EMPLOYEES);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      EMPLOYEE_HEADERS.length
    )
    .getValues();

  return values.map(function(row) {
    return {
      employeeCode: String(row[0] || ''),
      fullName: String(row[1] || ''),
      citizenID: normalizeCitizenID_(row[2]),
      department: String(row[3] || ''),
      section: String(row[4] || ''),
      position: String(row[5] || ''),
      mustChangePassword:
        row[8] === true ||
        String(row[8]).toLowerCase() === 'true',
      active:
        row[9] === true ||
        String(row[9]).toLowerCase() === 'true',
      updatedAt: String(row[10] || '')
    };
  });
}

function getLatestPayrollForEmployee_(
  employeeCode
) {
  const sheet =
    getSheet_(CONFIG.SHEETS.PAYROLL);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      PAYROLL_HEADERS.length
    )
    .getValues();

  let found = null;

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0]).trim() ===
      String(employeeCode).trim()
    ) {
      found = payrollRowToObject_(
        values[i]
      );
    }
  }

  return found;
}

function payrollRowToObject_(row) {
  const result = {};

  PAYROLL_HEADERS.forEach(function(header, index) {
    result[header] = row[index];
  });

  return result;
}

function createSession_(
  username,
  role
) {
  const token =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');

  const createdAt =
    new Date();

  const expiresAt =
    new Date(
      createdAt.getTime() +
      CONFIG.SESSION_SECONDS * 1000
    );

  const employeeCode =
    role === 'EMPLOYEE'
      ? username
      : '';

  const sheet =
    getSheet_(CONFIG.SHEETS.SESSIONS);

  sheet.appendRow([
    token,
    username,
    role,
    employeeCode,
    createdAt,
    expiresAt,
    true
  ]);

  return token;
}

function requireEmployeeSession_(token) {
  const session =
    findSession_(token);

  if (!session) {
    return null;
  }

  if (session.role !== 'EMPLOYEE') {
    return null;
  }

  return session;
}

function requireAdminSession_(token) {
  const session =
    findSession_(token);

  if (!session) {
    return null;
  }

  if (session.role !== 'ADMIN') {
    return null;
  }

  return session;
}

function findSession_(token) {
  if (!token) {
    return null;
  }

  const sheet =
    getSheet_(CONFIG.SHEETS.SESSIONS);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      SESSION_HEADERS.length
    )
    .getValues();

  const now =
    new Date();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0]) === token &&
      (
        values[i][6] === true ||
        String(values[i][6]).toLowerCase() === 'true'
      )
    ) {
      const expires =
        new Date(values[i][5]);

      if (
        isNaN(expires.getTime()) ||
        expires.getTime() <= now.getTime()
      ) {
        sheet.getRange(i + 2, 7).setValue(false);
        return null;
      }

      return {
        row: i + 2,
        token: token,
        username: String(values[i][1] || ''),
        role: String(values[i][2] || ''),
        employeeCode: String(values[i][3] || ''),
        createdAt: values[i][4],
        expiresAt: values[i][5]
      };
    }
  }

  return null;
}

function findAdmin_(username) {
  const sheet =
    getSheet_(CONFIG.SHEETS.ADMINS);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      ADMIN_HEADERS.length
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0]).trim() === username
    ) {
      return {
        row: i + 2,
        username: String(values[i][0]),
        passwordHash: String(values[i][1]),
        passwordSalt: String(values[i][2]),
        active:
          values[i][3] === true ||
          String(values[i][3]).toLowerCase() === 'true'
      };
    }
  }

  return null;
}

function createPasswordHash_(password) {
  const salt =
    Utilities.getUuid().replace(/-/g, '');

  const hash =
    hashPassword_(password, salt);

  return {
    salt: salt,
    hash: hash
  };
}

function hashPassword_(password, salt) {
  const input =
    String(salt) +
    ':' +
    String(password);

  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      input,
      Utilities.Charset.UTF_8
    );

  return digest
    .map(function(byte) {
      const value =
        byte < 0 ? byte + 256 : byte;

      return (
        value.toString(16).padStart(2, '0')
      );
    })
    .join('');
}

function verifyPassword_(
  password,
  hash,
  salt
) {
  if (!hash || !salt) {
    return false;
  }

  return (
    hashPassword_(password, salt) ===
    String(hash)
  );
}

function normalizeCitizenID_(value) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }

  let text =
    String(value).trim();

  if (
    /^\d+\.0+$/.test(text)
  ) {
    text =
      text.substring(
        0,
        text.indexOf('.')
      );
  }

  text =
    text.replace(/^'/, '');

  text =
    text.replace(/\s+/g, '');

  if (/^\d+$/.test(text)) {
    if (text.length < 12) {
      text =
        text.padStart(12, '0');
    }
  }

  return text;
}

function numberValue_(value) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return 0;
  }

  if (typeof value === 'number') {
    return value;
  }

  const text =
    String(value)
      .replace(/,/g, '')
      .trim();

  if (!text) {
    return 0;
  }

  const number =
    Number(text);

  return isNaN(number)
    ? 0
    : number;
}

function writeAudit_(
  username,
  role,
  action,
  target,
  ip,
  details
) {
  const sheet =
    getSheet_(CONFIG.SHEETS.AUDIT);

  sheet.appendRow([
    nowString_(),
    username,
    role,
    action,
    target,
    ip,
    details
  ]);
}

function getSheet_(name) {
  const ss =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(name);

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' + name
    );
  }

  return sheet;
}

function nowString_() {
  return Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    'dd/MM/yyyy HH:mm:ss'
  );
}

function parseRequest_(e) {
  if (!e || !e.postData) {
    return {};
  }

  const content =
    e.postData.contents || '';

  if (!content) {
    return {};
  }

  try {
    return JSON.parse(content);
  } catch (error) {
    throw new Error(
      'Invalid JSON request'
    );
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
