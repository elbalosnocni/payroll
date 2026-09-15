function sanitizePayrollForClient_(data) {
  return {
    employeeCode: data.EmployeeCode,
    fullName: data.FullName,
    cccd: data.CCCD,
    department: data.Department,
    section: data.Section,
    position: data.Position,

    baseSalary: data.BaseSalary,
    workingDays: data.WorkingDays,
    holidayDays: data.HolidayDays,
    paidLeaveDays: data.PaidLeaveDays,
    unpaidLeaveDays: data.UnpaidLeaveDays,

    overtimeHours: data.OvertimeHours,
    holidayWorkHours: data.HolidayWorkHours,
    holidayOvertimeHours: data.HolidayOvertimeHours,
    minimumRegionalPaidLeaveDays:
      data.MinimumRegionalPaidLeaveDays,
    holidayNightOvertimeHours:
      data.HolidayNightOvertimeHours,
    nightOvertimeHours: data.NightOvertimeHours,
    holidayHours: data.HolidayHours,
    holidayOvertimeHours:
      data.HolidayOvertimeHours,
    holidayNightOvertimeHours:
      data.HolidayNightOvertimeHours,
    nightShiftDays: data.NightShiftDays,

    otherIncome: data.OtherIncome,
    disciplinaryDeduction:
      data.DisciplinaryDeduction,

    loyalty2Years: data.Loyalty2Years,
    loyalty5Years: data.Loyalty5Years,
    loyalty10Years: data.Loyalty10Years,

    housingAllowance: data.HousingAllowance,
    transportationAllowance:
      data.TransportationAllowance,

    attendanceBonus:
      data.AttendanceBonus,

    terminationAndUnusedLeaveAllowance:
      data.TerminationAndUnusedLeaveAllowance,

    monthlySalary:
      data.MonthlySalary,

    overtimeHolidayNightSalary:
      data.OvertimeHolidayNightSalary,

    commissionAndExcessBonus:
      data.CommissionAndExcessBonus,

    otherIncomeTotal:
      data.OtherIncomeTotal,

    otherDeductions:
      data.OtherDeductions,

    advancePayment:
      data.AdvancePayment,

    socialInsurance:
      data.SocialInsurance,

    healthInsurance:
      data.HealthInsurance,

    unemploymentInsurance:
      data.UnemploymentInsurance,

    incomeTax:
      data.IncomeTax,

    netSalary:
      data.NetSalary,

    updatedAt:
      data.UpdatedAt
  };
}


function getPayrollForCurrentUser_(
  token,
  ipAddress
) {
  const payload =
    requireUserToken_(token);

  const employeeCode =
    normalizeEmployeeCode_(payload.sub);

  const employee =
    findEmployeeByCode_(employeeCode);

  if (!employee) {
    auditLog_(
      'USER',
      employeeCode,
      'GET_PAYROLL',
      employeeCode,
      false,
      ipAddress,
      'Employee not found'
    );

    throw new Error('EMPLOYEE_NOT_FOUND');
  }

  const payroll =
    sanitizePayrollForClient_(
      employee.data
    );

  auditLog_(
    'USER',
    employeeCode,
    'GET_PAYROLL',
    employeeCode,
    true,
    ipAddress,
    'Payroll returned'
  );

  return createResponse_(
    true,
    payroll,
    'Lấy dữ liệu lương thành công.',
    'OK'
  );
}


function validateSyncKey_(providedKey) {
  return timingSafeEqual_(
    String(providedKey || ''),
    getSyncKey_()
  );
}


function normalizeSyncRecord_(record) {
  const object = {};

  CONFIG.DATA_HEADERS.forEach(function(header) {
    object[header] = '';
  });

  CONFIG.DATA_HEADERS.forEach(function(header) {
    if (record[header] !== undefined) {
      object[header] = record[header];
    }
  });

  object.EmployeeCode =
    normalizeEmployeeCode_(
      object.EmployeeCode
    );

  object.FullName =
    normalizeString_(
      object.FullName
    );

  object.CCCD =
    normalizeCCCD_(
      object.CCCD
    );

  object.Department =
    normalizeString_(
      object.Department
    );

  object.Section =
    normalizeString_(
      object.Section
    );

  object.Position =
    normalizeString_(
      object.Position
    );

  object.PasswordHash =
    normalizeString_(
      object.PasswordHash
    );

  object.PasswordSalt =
    normalizeString_(
      object.PasswordSalt
    );

  object.MustChangePassword =
    object.MustChangePassword === true ||
    String(object.MustChangePassword)
      .toLowerCase() === 'true';

  object.UpdatedAt =
    normalizeString_(
      object.UpdatedAt
    ) ||
    formatDateTime_(new Date());

  object.UpdatedBy =
    normalizeString_(
      object.UpdatedBy
    ) ||
    'VBA_SYNC';

  return object;
}


function syncPayrollData_(
  syncKey,
  records,
  sourceFile,
  sourceMonth,
  ipAddress
) {
  if (!validateSyncKey_(syncKey)) {
    auditLog_(
      'SYSTEM',
      'VBA',
      'SYNC',
      sourceFile || '',
      false,
      ipAddress,
      'Invalid sync key'
    );

    throw new Error('INVALID_SYNC_KEY');
  }

  if (!Array.isArray(records)) {
    throw new Error('INVALID_SYNC_RECORDS');
  }

  if (
    records.length === 0 ||
    records.length > CONFIG.SYNC.MAX_ROWS_PER_REQUEST
  ) {
    throw new Error('INVALID_SYNC_RECORD_COUNT');
  }

  const lock =
    acquireScriptLock_(30000);

  try {
    const sheet =
      getSheet_(CONFIG.SHEETS.DATA);

    const headers =
      getHeaders_(sheet);

    const employeeCodeColumn =
      findColumnIndex_(
        headers,
        'EmployeeCode'
      );

    const cccdColumn =
      findColumnIndex_(
        headers,
        'CCCD'
      );

    const existingRowMap =
      {};

    if (sheet.getLastRow() >= 2) {
      const existingValues =
        sheet
          .getRange(
            2,
            employeeCodeColumn + 1,
            sheet.getLastRow() - 1,
            1
          )
          .getDisplayValues();

      existingValues.forEach(
        function(row, index) {
          const code =
            normalizeEmployeeCode_(row[0]);

          if (code) {
            existingRowMap[code] =
              index + 2;
          }
        }
      );
    }

    const inserts = [];
    const updates = [];

    records.forEach(function(record) {
      const normalized =
        normalizeSyncRecord_(record);

      if (!normalized.EmployeeCode) {
        throw new Error(
          'SYNC_RECORD_MISSING_EMPLOYEE_CODE'
        );
      }

      if (!normalized.FullName) {
        throw new Error(
          'SYNC_RECORD_MISSING_FULL_NAME'
        );
      }

      if (!isValidCCCD_(normalized.CCCD)) {
        throw new Error(
          'SYNC_RECORD_INVALID_CCCD_' +
          normalized.EmployeeCode
        );
      }

      const row =
        objectToRow_(
          headers,
          normalized
        );

      if (
        existingRowMap[
          normalized.EmployeeCode
        ]
      ) {
        updates.push({
          rowNumber:
            existingRowMap[
              normalized.EmployeeCode
            ],
          row: row
        });
      } else {
        inserts.push(row);
      }
    });

    updates.forEach(function(item) {
      sheet
        .getRange(
          item.rowNumber,
          1,
          1,
          headers.length
        )
        .setValues([item.row]);
    });

    if (inserts.length > 0) {
      const startRow =
        sheet.getLastRow() + 1;

      sheet
        .getRange(
          startRow,
          1,
          inserts.length,
          headers.length
        )
        .setValues(inserts);
    }

    const cccdRange =
      sheet.getRange(
        2,
        cccdColumn + 1,
        Math.max(
          sheet.getMaxRows() - 1,
          1
        ),
        1
      );

    cccdRange.setNumberFormat('@');

    auditLog_(
      'SYSTEM',
      'VBA',
      'SYNC',
      sourceFile || '',
      true,
      ipAddress,
      JSON.stringify({
        sourceMonth: sourceMonth || '',
        total: records.length,
        updated: updates.length,
        inserted: inserts.length
      })
    );

    return createResponse_(
      true,
      {
        received: records.length,
        updated: updates.length,
        inserted: inserts.length,
        sourceFile: sourceFile || '',
        sourceMonth: sourceMonth || '',
        updatedAt: formatDateTime_(new Date())
      },
      'Đồng bộ dữ liệu thành công.',
      'OK'
    );
  } finally {
    lock.releaseLock();
  }
}
