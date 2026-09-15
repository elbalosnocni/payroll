const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1Y-f3xSzwrV2ycBnzThlrtgcOcHD2Z4iqLKqqTyVFejM',

  SHEETS: Object.freeze({
    DATA: 'Data',
    ADMIN: 'Admin',
    AUDIT: 'AuditLog'
  }),

  API: Object.freeze({
    TOKEN_TTL_SECONDS: 30 * 60,
    MAX_LOGIN_ATTEMPTS: 5,
    LOGIN_LOCK_SECONDS: 15 * 60,
    MAX_BODY_BYTES: 45 * 1024 * 1024
  }),

  PASSWORD: Object.freeze({
    ITERATIONS: 12000,
    SALT_BYTES: 16,
    MIN_LENGTH: 8,
    MAX_LENGTH: 128
  }),

  SESSION: Object.freeze({
    TOKEN_BYTES: 32
  }),

  SYNC: Object.freeze({
    MAX_ROWS_PER_REQUEST: 10000
  }),

  PROPERTIES: Object.freeze({
    SECRET_KEY: 'SALARY_API_SECRET_KEY',
    SYNC_KEY: 'SALARY_SYNC_KEY'
  }),

  DATA_HEADERS: Object.freeze([
    'EmployeeCode',
    'FullName',
    'CCCD',
    'Department',
    'Section',
    'Position',

    'BaseSalary',
    'WorkingDays',
    'HolidayDays',
    'PaidLeaveDays',
    'UnpaidLeaveDays',

    'OvertimeHours',
    'HolidayWorkHours',
    'HolidayOvertimeHours',
    'MinimumRegionalPaidLeaveDays',
    'HolidayNightOvertimeHours',
    'NightOvertimeHours',
    'HolidayHours',
    'HolidayOvertimeHours',
    'HolidayNightOvertimeHours',
    'NightShiftDays',

    'OtherIncome',
    'DisciplinaryDeduction',
    'Loyalty2Years',
    'Loyalty5Years',
    'Loyalty10Years',
    'HousingAllowance',
    'TransportationAllowance',
    'AttendanceBonus',
    'TerminationAndUnusedLeaveAllowance',

    'MonthlySalary',
    'OvertimeHolidayNightSalary',
    'CommissionAndExcessBonus',
    'OtherIncomeTotal',

    'OtherDeductions',
    'AdvancePayment',
    'SocialInsurance',
    'HealthInsurance',
    'UnemploymentInsurance',
    'IncomeTax',

    'NetSalary',

    'PasswordHash',
    'PasswordSalt',
    'MustChangePassword',

    'UpdatedAt',
    'UpdatedBy'
  ]),

  ADMIN_HEADERS: Object.freeze([
    'AdminId',
    'Username',
    'PasswordHash',
    'PasswordSalt',
    'Active',
    'CreatedAt',
    'UpdatedAt'
  ]),

  AUDIT_HEADERS: Object.freeze([
    'Timestamp',
    'ActorType',
    'ActorId',
    'Action',
    'Target',
    'Success',
    'IPAddress',
    'Details'
  ])
});


function getSpreadsheet_() {
  const id = String(CONFIG.SPREADSHEET_ID || '').trim();

  if (!id || id === 'PUT_YOUR_GOOGLE_SHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID chưa được cấu hình.');
  }

  return SpreadsheetApp.openById(id);
}


function getSheet_(sheetName) {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + sheetName);
  }

  return sheet;
}


function initializeConfiguration() {
  const properties = PropertiesService.getScriptProperties();

  let secretKey = properties.getProperty(CONFIG.PROPERTIES.SECRET_KEY);
  let syncKey = properties.getProperty(CONFIG.PROPERTIES.SYNC_KEY);

  if (!secretKey) {
    secretKey = randomHex_(64);
    properties.setProperty(CONFIG.PROPERTIES.SECRET_KEY, secretKey);
  }

  if (!syncKey) {
    syncKey = randomHex_(64);
    properties.setProperty(CONFIG.PROPERTIES.SYNC_KEY, syncKey);
  }

  initializeSheets_();

  return {
    success: true,
    message: 'Configuration initialized.',
    secretKeyCreated: true,
    syncKeyCreated: true
  };
}


function initializeSheets_() {
  const spreadsheet = getSpreadsheet_();

  let dataSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DATA);
  let adminSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ADMIN);
  let auditSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.AUDIT);

  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(CONFIG.SHEETS.DATA);
  }

  if (!adminSheet) {
    adminSheet = spreadsheet.insertSheet(CONFIG.SHEETS.ADMIN);
  }

  if (!auditSheet) {
    auditSheet = spreadsheet.insertSheet(CONFIG.SHEETS.AUDIT);
  }

  ensureHeaderRow_(
    dataSheet,
    CONFIG.DATA_HEADERS
  );

  ensureHeaderRow_(
    adminSheet,
    CONFIG.ADMIN_HEADERS
  );

  ensureHeaderRow_(
    auditSheet,
    CONFIG.AUDIT_HEADERS
  );

  const cccdIndex = CONFIG.DATA_HEADERS.indexOf('CCCD') + 1;

  if (cccdIndex > 0) {
    dataSheet
      .getRange(2, cccdIndex, Math.max(dataSheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat('@');
  }

  return true;
}


function ensureHeaderRow_(sheet, headers) {
  const currentColumnCount = sheet.getMaxColumns();

  if (currentColumnCount < headers.length) {
    sheet.insertColumnsAfter(
      currentColumnCount,
      headers.length - currentColumnCount
    );
  }

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight('bold');
}


function getSecretKey_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.SECRET_KEY);

  if (!value) {
    throw new Error(
      'Secret key chưa được thiết lập. Hãy chạy initializeConfiguration().'
    );
  }

  return value;
}


function getSyncKey_() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.SYNC_KEY);

  if (!value) {
    throw new Error(
      'Sync key chưa được thiết lập. Hãy chạy initializeConfiguration().'
    );
  }

  return value;
}


function getApiOrigin_() {
  return 'https://script.google.com';
}
