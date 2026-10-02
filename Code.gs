const CONFIG = {
  SPREADSHEET_ID: '1Y-f3xSzwrV2ycBnzThlrtgcOcHD2Z4iqLKqqTyVFejM',

  SESSION_SECONDS: 3600,

  LOGIN_MAX_ATTEMPTS: 5,
  LOGIN_LOCK_MINUTES: 15,
  SESSION_CACHE_SECONDS: 300,
  ADMIN_DEFAULT_ROLE: 'SUPER_ADMIN',

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
    SYNC: 'SyncStatus',
    PAYROLL_CONFIG: 'PayrollConfig',
    SYSTEM_SETTINGS: 'SystemSettings'
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
  'UpdatedAt',
  'FailedAttempts',
  'LockedUntil',
  'LastLoginAt'
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
  'NightHolidayOvertimeDayHours',

  'UpdatedAt',
  'ExtraData'
];

const ADMIN_HEADERS = [
  'Username',
  'PasswordHash',
  'PasswordSalt',
  'Active',
  'UpdatedAt',
  'Role',
  'FailedAttempts',
  'LockedUntil',
  'MustChangePassword',
  'LastLoginAt',
  'LastPasswordChange'
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

const PAYROLL_CONFIG_HEADERS = [
  'FieldKey', 'Label', 'Type', 'Group', 'Subgroup', 'Visible', 'Order', 'Active'
];

const SYSTEM_SETTINGS_HEADERS = [
  'Key', 'Value', 'UpdatedAt'
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
    version: '2.0.0',
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

      case 'adminDashboard':
        return jsonResponse_(adminDashboard_(body));

      case 'adminSetEmployeeStatus':
        return jsonResponse_(adminSetEmployeeStatus_(body));

      case 'adminChangePassword':
        return jsonResponse_(adminChangePassword_(body));

      case 'adminListAdmins':
        return jsonResponse_(adminListAdmins_(body));

      case 'adminCreateAdmin':
        return jsonResponse_(adminCreateAdmin_(body));

      case 'adminSetAdminStatus':
        return jsonResponse_(adminSetAdminStatus_(body));

      case 'adminPayrollConfig':
        return jsonResponse_(adminPayrollConfig_(body));

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
  if (String(body.setupKey || '') !== getSetting_('SETUP_KEY', '1235678')) {
    return { success: false, error: 'INVALID_SETUP_KEY' };
  }
  ensureSystem_();
  if (!getSetting_('SYNC_API_KEY','')) setSetting_('SYNC_API_KEY', CONFIG.SYNC_API_KEY);
  if (!getSetting_('SETUP_KEY','')) setSetting_('SETUP_KEY', '1235678');
  const sheet = getSheet_(CONFIG.SHEETS.ADMINS);
  const admin = findAdmin_(CONFIG.DEFAULT_ADMIN_USERNAME);
  if (!admin) {
    const passwordData = createPasswordHash_(CONFIG.DEFAULT_ADMIN_PASSWORD);
    sheet.appendRow([CONFIG.DEFAULT_ADMIN_USERNAME, passwordData.hash, passwordData.salt, true, nowString_(), CONFIG.ADMIN_DEFAULT_ROLE, 0, '', true, '', nowString_()]);
  }
  writeAudit_(CONFIG.DEFAULT_ADMIN_USERNAME, CONFIG.ADMIN_DEFAULT_ROLE, 'SYSTEM_SETUP', '', '', 'System initialized / migrated');
  return { success: true, message: 'SYSTEM_READY', adminUsername: CONFIG.DEFAULT_ADMIN_USERNAME };
}

function ensureSystem_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureSheet_(ss, CONFIG.SHEETS.EMPLOYEES, EMPLOYEE_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.PAYROLL, PAYROLL_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, SESSION_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.AUDIT, AUDIT_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.SYNC, SYNC_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.ADMINS, ADMIN_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.PAYROLL_CONFIG, PAYROLL_CONFIG_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.SYSTEM_SETTINGS, SYSTEM_SETTINGS_HEADERS);
  seedPayrollConfig_();
  const adminSheet = getSheet_(CONFIG.SHEETS.ADMINS);
  if (!findAdmin_(CONFIG.DEFAULT_ADMIN_USERNAME)) {
    const passwordData = createPasswordHash_(CONFIG.DEFAULT_ADMIN_PASSWORD);
    adminSheet.appendRow([CONFIG.DEFAULT_ADMIN_USERNAME, passwordData.hash, passwordData.salt, true, nowString_(), CONFIG.ADMIN_DEFAULT_ROLE, 0, '', true, '', nowString_()]);
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  headers.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      current.push(header);
    }
  });
  sheet.setFrozenRows(1);
  return sheet;
}

function seedPayrollConfig_() {
  const sheet = getSheet_(CONFIG.SHEETS.PAYROLL_CONFIG);
  const lastRow = sheet.getLastRow();
  const existing = {};
  if (lastRow > 1) {
    sheet.getRange(2,1,lastRow-1,PAYROLL_CONFIG_HEADERS.length).getValues().forEach(r => existing[String(r[0]||'').trim()] = true);
  }
  const labels = { BasicSalary:'Lương cơ bản', WorkingDays:'Số ngày làm việc', HolidayDays:'Số ngày lễ', PaidLeaveDays:'Số ngày nghỉ hưởng lương', UnpaidLeaveDays:'Số ngày nghỉ không hưởng lương', OvertimeHours:'Số giờ làm ngoài giờ', HolidayWorkHours:'Số giờ làm ngày nghỉ', HolidayOvertimeHours:'Số giờ làm ngoài giờ ngày nghỉ', OtherMoney:'Tiền khác', HousingAllowance:'Tiền nhà ở', TransportationAllowance:'Tiền đi lại', AttendanceBonus:'Tiền thưởng chuyên cần', SocialInsurance:'BHXH', HealthInsurance:'BHYT', UnemploymentInsurance:'BHTN', PersonalIncomeTax:'Thuế thu nhập', AdvancePayment:'Tạm ứng', OtherDeductions:'Các khoản khấu trừ khác', GrossIncome:'Tổng thu nhập', NetSalary:'Lương thực lĩnh' };
  const rows=[];
  PAYROLL_HEADERS.forEach(function(h,i){
    if (h==='ExtraData' || existing[h]) return;
    rows.push([h, labels[h] || h, /Salary|Income|Money|Allowance|Bonus|Insurance|Tax|Payment|Deductions|Loyalty|Severance|Commission/.test(h) ? 'money' : 'number', '', '', true, i+1, true]);
  });
  if (rows.length) sheet.getRange(sheet.getLastRow()+1,1,rows.length,PAYROLL_CONFIG_HEADERS.length).setValues(rows);
}

function getSetting_(key, fallback) {
  try {
    const s=getSheet_(CONFIG.SHEETS.SYSTEM_SETTINGS), last=s.getLastRow();
    if(last>1){ const rows=s.getRange(2,1,last-1,3).getValues(); for(const r of rows) if(String(r[0])===key) return String(r[1]||fallback); }
  } catch(e) {}
  return fallback;
}

function setSetting_(key, value) {
  const s=getSheet_(CONFIG.SHEETS.SYSTEM_SETTINGS), last=s.getLastRow();
  if(last>1){ const rows=s.getRange(2,1,last-1,3).getValues(); for(let i=0;i<rows.length;i++){ if(String(rows[i][0])===key){ s.getRange(i+2,2,1,2).setValues([[String(value),nowString_()]]); return; } } }
  s.appendRow([key,String(value),nowString_()]);
}

function login_(body) {
  ensureSystem_(); const citizenID=normalizeCitizenID_(body.citizenID), password=String(body.password||''); if(!citizenID||!password)return {success:false,error:'INVALID_LOGIN'};
  const employee=findEmployeeByCitizenID_(citizenID); if(!employee||!employee.active)return {success:false,error:'INVALID_LOGIN'};
  if(isLockedUntil_(employee.lockedUntil))return {success:false,error:'ACCOUNT_LOCKED'};
  if(!verifyPassword_(password,employee.passwordHash,employee.passwordSalt)){
    const attempts=(employee.failedAttempts||0)+1; const sh=getSheet_(CONFIG.SHEETS.EMPLOYEES); sh.getRange(employee.row,12).setValue(attempts); if(attempts>=CONFIG.LOGIN_MAX_ATTEMPTS)sh.getRange(employee.row,13).setValue(new Date(Date.now()+CONFIG.LOGIN_LOCK_MINUTES*60000));
    writeAudit_(employee.employeeCode,'EMPLOYEE','LOGIN_FAILED',employee.employeeCode,'','Invalid password; attempts='+attempts); return {success:false,error:attempts>=CONFIG.LOGIN_MAX_ATTEMPTS?'ACCOUNT_LOCKED':'INVALID_LOGIN'};
  }
  const sh=getSheet_(CONFIG.SHEETS.EMPLOYEES); sh.getRange(employee.row,12,1,3).setValues([[0,'',nowString_()]]); const token=createSession_(employee.employeeCode,'EMPLOYEE');
  writeAudit_(employee.employeeCode,'EMPLOYEE','LOGIN',employee.employeeCode,'','Employee login'); const response={success:true,token:token,mustChangePassword:employee.mustChangePassword,employee:{employeeCode:employee.employeeCode,fullName:employee.fullName,citizenID:employee.citizenID,department:employee.department,section:employee.section,position:employee.position}};
  if(!employee.mustChangePassword){const payrolls=getPayrollHistoryForEmployee_(employee.employeeCode); response.payrolls=payrolls; response.payroll=payrolls.length?payrolls[0]:null; response.payrollConfig=readPayrollConfig_();} return response;
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

  const payrolls = getPayrollHistoryForEmployee_(
    employee.employeeCode
  );

  return {
    success: true,
    payrolls: payrolls,
    payroll: payrolls.length ? payrolls[0] : null,
    payrollConfig: readPayrollConfig_()
  };
}

function logout_(body) {
  const token = String(body.token || '');

  if (!token) {
    return {
      success: true
    };
  }

  CacheService.getScriptCache().remove('sess:' + token);
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
  const username=String(body.username||'').trim(); const password=String(body.password||'');
  const admin=findAdmin_(username);
  if(!admin || !admin.active) { writeAudit_(username,'ADMIN','ADMIN_LOGIN_FAILED',username,'','Admin not found/inactive'); return {success:false,error:'INVALID_LOGIN'}; }
  if(isLockedUntil_(admin.lockedUntil)) return {success:false,error:'ACCOUNT_LOCKED'};
  if(!verifyPassword_(password,admin.passwordHash,admin.passwordSalt)){
    const attempts=(Number(admin.failedAttempts)||0)+1;
    const sheet=getSheet_(CONFIG.SHEETS.ADMINS);
    if(attempts>=CONFIG.LOGIN_MAX_ATTEMPTS){ sheet.getRange(admin.row,8).setValue(nowString_()); sheet.getRange(admin.row,7).setValue(attempts); sheet.getRange(admin.row,8).setValue(new Date(Date.now()+CONFIG.LOGIN_LOCK_MINUTES*60000)); } else sheet.getRange(admin.row,7).setValue(attempts);
    writeAudit_(username,'ADMIN','ADMIN_LOGIN_FAILED',username,'','Invalid password; attempts='+attempts);
    return {success:false,error:attempts>=CONFIG.LOGIN_MAX_ATTEMPTS?'ACCOUNT_LOCKED':'INVALID_LOGIN'};
  }
  const sheet=getSheet_(CONFIG.SHEETS.ADMINS); sheet.getRange(admin.row,7).setValue(0); sheet.getRange(admin.row,8).setValue(''); sheet.getRange(admin.row,10).setValue(nowString_());
  const role=admin.role||CONFIG.ADMIN_DEFAULT_ROLE; const token=createSession_(username,role);
  writeAudit_(username,role,'ADMIN_LOGIN',username,'','Admin login');
  return {success:true,token:token,role:role,username:username,mustChangePassword:admin.mustChangePassword};
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
  const session=requireAdminSession_(body.token); if(!session) return {success:false,error:'UNAUTHORIZED'};
  const employeeCode=String(body.employeeCode||'').trim(); if(!employeeCode) return {success:false,error:'EMPLOYEE_CODE_REQUIRED'};
  const employee=findEmployeeByEmployeeCode_(employeeCode); if(!employee) return {success:false,error:'EMPLOYEE_NOT_FOUND'};
  const pd=createPasswordHash_(employee.employeeCode); const sheet=getSheet_(CONFIG.SHEETS.EMPLOYEES);
  sheet.getRange(employee.row,7,1,4).setValues([[pd.hash,pd.salt,true,true]]);
  sheet.getRange(employee.row,12,1,2).setValues([[0,'']]); sheet.getRange(employee.row,11).setValue(nowString_());
  writeAudit_(session.username,session.role,'RESET_PASSWORD',employeeCode,'','Password reset to employee code');
  return {success:true,message:'PASSWORD_RESET',employeeCode:employeeCode};
}

function adminDashboard_(body) {
  const session=requireAdminSession_(body.token); if(!session) return {success:false,error:'UNAUTHORIZED'};
  const employees=readEmployees_(); const admins=readAdmins_();
  const activeEmployees=employees.filter(e=>e.active).length; const lockedEmployees=employees.filter(e=>isLockedUntil_(e.lockedUntil)).length;
  const mustChange=employees.filter(e=>e.mustChangePassword).length;
  const sync=adminSyncStatus_({token:body.token});
  return {success:true,stats:{employees:employees.length,activeEmployees:activeEmployees,lockedEmployees:lockedEmployees,mustChangePassword:mustChange,admins:admins.length,activeAdmins:admins.filter(a=>a.active).length,payrollRows:countPayrollRows_()},sync:sync.status||null};
}

function adminSetEmployeeStatus_(body) {
  const session=requireAdminSession_(body.token); if(!session) return {success:false,error:'UNAUTHORIZED'};
  const code=String(body.employeeCode||'').trim(); const active=body.active===true || String(body.active).toLowerCase()==='true';
  const emp=findEmployeeByEmployeeCode_(code); if(!emp) return {success:false,error:'EMPLOYEE_NOT_FOUND'};
  getSheet_(CONFIG.SHEETS.EMPLOYEES).getRange(emp.row,10).setValue(active); getSheet_(CONFIG.SHEETS.EMPLOYEES).getRange(emp.row,11).setValue(nowString_());
  writeAudit_(session.username,session.role,active?'UNLOCK_EMPLOYEE':'LOCK_EMPLOYEE',code,'','Employee status changed');
  return {success:true,active:active};
}

function adminChangePassword_(body) {
  const session=requireAdminSession_(body.token); if(!session) return {success:false,error:'UNAUTHORIZED'};
  const current=String(body.currentPassword||''), next=String(body.newPassword||''); if(next.length<8) return {success:false,error:'PASSWORD_TOO_SHORT'};
  const admin=findAdmin_(session.username); if(!admin || !verifyPassword_(current,admin.passwordHash,admin.passwordSalt)) return {success:false,error:'CURRENT_PASSWORD_INVALID'};
  const pd=createPasswordHash_(next); const s=getSheet_(CONFIG.SHEETS.ADMINS); s.getRange(admin.row,2,1,2).setValues([[pd.hash,pd.salt]]); s.getRange(admin.row,9).setValue(false); s.getRange(admin.row,11).setValue(nowString_()); s.getRange(admin.row,5).setValue(nowString_());
  writeAudit_(session.username,session.role,'ADMIN_CHANGE_PASSWORD',session.username,'','Admin changed own password');
  return {success:true};
}

function readAdmins_(){
  const s=getSheet_(CONFIG.SHEETS.ADMINS), last=s.getLastRow(); if(last<=1)return[];
  return s.getRange(2,1,last-1,ADMIN_HEADERS.length).getValues().map((r,i)=>({row:i+2,username:String(r[0]||''),active:bool_(r[3]),role:String(r[5]||CONFIG.ADMIN_DEFAULT_ROLE),failedAttempts:Number(r[6]||0),lockedUntil:r[7]||'',mustChangePassword:bool_(r[8]),lastLoginAt:r[9]||''}));
}
function adminListAdmins_(body){ const session=requireAdminSession_(body.token); if(!session||!isSuperAdmin_(session))return {success:false,error:'FORBIDDEN'}; return {success:true,admins:readAdmins_()}; }
function adminCreateAdmin_(body){
  const session=requireAdminSession_(body.token); if(!session||!isSuperAdmin_(session))return {success:false,error:'FORBIDDEN'};
  const username=String(body.username||'').trim(); const password=String(body.password||''); const role=String(body.role||'HR_ADMIN').trim();
  if(!/^[A-Za-z0-9._-]{3,50}$/.test(username))return {success:false,error:'INVALID_USERNAME'}; if(password.length<8)return {success:false,error:'PASSWORD_TOO_SHORT'}; if(findAdmin_(username))return {success:false,error:'ADMIN_EXISTS'};
  const pd=createPasswordHash_(password); getSheet_(CONFIG.SHEETS.ADMINS).appendRow([username,pd.hash,pd.salt,true,nowString_(),role,0,'',false,'',nowString_()]);
  writeAudit_(session.username,session.role,'CREATE_ADMIN',username,'','Admin created role='+role); return {success:true};
}
function adminSetAdminStatus_(body){
  const session=requireAdminSession_(body.token); if(!session||!isSuperAdmin_(session))return {success:false,error:'FORBIDDEN'};
  const username=String(body.username||'').trim(); const active=body.active===true || String(body.active).toLowerCase()==='true'; if(username===session.username && !active)return {success:false,error:'CANNOT_DISABLE_SELF'};
  const a=findAdmin_(username); if(!a)return {success:false,error:'ADMIN_NOT_FOUND'}; getSheet_(CONFIG.SHEETS.ADMINS).getRange(a.row,4).setValue(active); getSheet_(CONFIG.SHEETS.ADMINS).getRange(a.row,5).setValue(nowString_());
  writeAudit_(session.username,session.role,active?'UNLOCK_ADMIN':'LOCK_ADMIN',username,'','Admin status changed'); return {success:true};
}
function adminPayrollConfig_(body){
  const session=requireAdminSession_(body.token); if(!session)return {success:false,error:'UNAUTHORIZED'};
  return {success:true,config:readPayrollConfig_()};
}
function readPayrollConfig_(){
  const s=getSheet_(CONFIG.SHEETS.PAYROLL_CONFIG), last=s.getLastRow(); if(last<=1)return[];
  return s.getRange(2,1,last-1,PAYROLL_CONFIG_HEADERS.length).getValues().map(r=>({fieldKey:String(r[0]||''),label:String(r[1]||r[0]||''),type:String(r[2]||'number'),group:String(r[3]||''),subgroup:String(r[4]||''),visible:bool_(r[5]),order:Number(r[6]||0),active:bool_(r[7])})).filter(x=>x.active);
}
function countPayrollRows_(){ const s=getSheet_(CONFIG.SHEETS.PAYROLL); return Math.max(0,s.getLastRow()-1); }
function isSuperAdmin_(session){ return String(session.role||'').toUpperCase()==='SUPER_ADMIN'; }
function bool_(v){return v===true || String(v).toLowerCase()==='true' || String(v)==='1';}
function isLockedUntil_(v){ if(!v)return false; const d=v instanceof Date?v:new Date(v); return !isNaN(d.getTime()) && d.getTime()>Date.now(); }

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
      month: normalizeSalaryMonth_(row[2]),
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
    getSetting_('SYNC_API_KEY', CONFIG.SYNC_API_KEY)
  ) {
    return {
      success: false,
      error: 'INVALID_SYNC_KEY'
    };
  }

  ensureSystem_();

  const month = normalizeSalaryMonth_(body.salaryMonth);

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

    // Enforce one payslip per EmployeeCode + SalaryMonth.
    // If the same employee appears more than once in the incoming files,
    // the last record is used instead of creating duplicate payslips.
    const uniqueRecords = {};
    records.forEach(function(record) {
      const code = String(record.employeeCode || '').trim();
      if (code) uniqueRecords[code.toUpperCase()] = record;
    });

    let employeeCount = 0;

    Object.keys(uniqueRecords).forEach(function(codeKey) {
      const record = uniqueRecords[codeKey];
      const employeeCode =
        String(record.employeeCode || '').trim();

      const fullName =
        String(record.fullName || '').trim();

      const citizenID =
        normalizeCitizenID_(record.citizenID);

      if (!employeeCode || !fullName) {
        return;
      }

      if(record.extraData && typeof record.extraData !== 'object') record.extraData = {};
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

    const replacedCount = deletePayrollMonth_(month);

    if (payrollRows.length > 0) {
      const startRow = payrollSheet.getLastRow() + 1;
      // Keep SalaryMonth as text so Google Sheets cannot turn 08-2026 into a Date.
      payrollSheet.getRange(2, 7, Math.max(1, payrollSheet.getMaxRows() - 1), 1).setNumberFormat('@');
      payrollSheet
        .getRange(startRow, 1, payrollRows.length, PAYROLL_HEADERS.length)
        .setValues(payrollRows);
    }

    const statusSheet =
      getSheet_(CONFIG.SHEETS.SYNC);

    const statusRow = statusSheet.getLastRow() + 1;
    statusSheet.getRange(statusRow, 3).setNumberFormat('@');
    statusSheet.getRange(statusRow, 1, 1, SYNC_HEADERS.length).setValues([[
      nowString_(),
      'SUCCESS',
      month,
      source.snackFile || '',
      source.flexibleFile || '',
      employeeCount,
      payrollRows.length,
      'Payroll synchronized successfully'
    ]]);

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
      payrollCount: payrollRows.length,
      replacedCount: replacedCount
    };
  } catch (error) {
    const statusSheet =
      getSheet_(CONFIG.SHEETS.SYNC);

    const errorRow = statusSheet.getLastRow() + 1;
    statusSheet.getRange(errorRow, 3).setNumberFormat('@');
    statusSheet.getRange(errorRow, 1, 1, SYNC_HEADERS.length).setValues([[
      nowString_(),
      'ERROR',
      month,
      source.snackFile || '',
      source.flexibleFile || '',
      0,
      0,
      String(error.message || error)
    ]]);

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
  // Excel khong co cot OtherIncome rieng. Muc III la tong cac subgroup.
  const otherIncome =
    numberValue_(record.otherMoney) +
    numberValue_(record.disciplinaryMoney) +
    numberValue_(record.loyalty2Years) +
    numberValue_(record.loyalty5Years) +
    numberValue_(record.loyalty10Years) +
    numberValue_(record.housingAllowance) +
    numberValue_(record.transportationAllowance) +
    numberValue_(record.attendanceBonus) +
    numberValue_(record.commissionAndOverTargetBonus) +
    numberValue_(record.severanceAndUnusedLeave);

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
    otherIncome,

    numberValue_(record.otherDeductions),
    numberValue_(record.advancePayment),

    numberValue_(record.socialInsurance),
    numberValue_(record.healthInsurance),
    numberValue_(record.unemploymentInsurance),
    numberValue_(record.personalIncomeTax),

    numberValue_(record.grossIncome),
    numberValue_(record.netSalary),
    numberValue_(record.nightHolidayOvertimeDayHours),

    nowString_(),
    JSON.stringify(record.extraData || {})
  ];
}

function deletePayrollMonth_(month) {
  const sheet = getSheet_(CONFIG.SHEETS.PAYROLL);
  const targetMonth = normalizeSalaryMonth_(month);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1 || !targetMonth) {
    return 0;
  }

  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, PAYROLL_HEADERS.length).getValues();
  const kept = [];
  let removed = 0;

  values.forEach(function(row) {
    if (normalizeSalaryMonth_(row[6]) === targetMonth) {
      removed++;
    } else {
      kept.push(row);
    }
  });

  // Do not use deleteRow(). It can fail when the sheet has frozen rows.
  // Clear the data area and rewrite only the rows that must remain.
  sheet.getRange(2, 1, rowCount, PAYROLL_HEADERS.length).clearContent();

  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, PAYROLL_HEADERS.length).setValues(kept);
  }

  return removed;
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
      employee.employeeCode, employee.fullName, employee.citizenID, employee.department, employee.section, employee.position,
      passwordData.hash, passwordData.salt, true, true, nowString_(), 0, '', ''
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
    updatedAt: String(row[10] || ''),
    failedAttempts: Number(row[11] || 0),
    lockedUntil: row[12] || '',
    lastLoginAt: row[13] || ''
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
      updatedAt: String(row[10] || ''),
      failedAttempts: Number(row[11] || 0),
      lockedUntil: row[12] || '',
      lastLoginAt: row[13] || ''
    };
  });
}

function getPayrollHistoryForEmployee_(employeeCode) {
  const sheet = getSheet_(CONFIG.SHEETS.PAYROLL);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, PAYROLL_HEADERS.length).getValues();
  const result = [];
  const seen = new Set();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() !== String(employeeCode || '').trim()) {
      continue;
    }

    const item = payrollRowToObject_(values[i]);
    const month = normalizeSalaryMonth_(item.SalaryMonth);
    const key = String(employeeCode || '').trim().toUpperCase() + '|' + month;

    // One payslip per employee per month. If old duplicate rows exist,
    // keep only one in the history response.
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }

  result.sort(function(a, b) {
    return salaryMonthSortKey_(b.SalaryMonth) - salaryMonthSortKey_(a.SalaryMonth);
  });

  return result;
}

function getLatestPayrollForEmployee_(employeeCode) {
  const history = getPayrollHistoryForEmployee_(employeeCode);
  return history.length ? history[0] : null;
}

function salaryMonthSortKey_(value) {
  const m = normalizeSalaryMonth_(value).match(/^(\d{2})-(\d{4})$/);
  if (!m) return 0;
  return Number(m[2]) * 100 + Number(m[1]);
}

function normalizeSalaryMonth_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  // Google Sheets may return a cell formatted as a date as a JavaScript Date.
  // Example: 2026-07-31T17:00:00.000Z is 2026-08-01 00:00 in Vietnam.
  // For payroll, always display/store the business month as MM-YYYY.
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh',
      'MM-yyyy'
    );
  }

  const text = String(value).trim();

  // Already correct.
  if (/^\d{2}-\d{4}$/.test(text)) {
    return text;
  }

  // ISO date/time returned by an API or sheet value.
  const iso = new Date(text);
  if (!isNaN(iso.getTime()) && /T|Z|\d{4}-\d{2}-\d{2}/.test(text)) {
    return Utilities.formatDate(
      iso,
      Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh',
      'MM-yyyy'
    );
  }

  // Accept common textual forms such as 8-2026 / 08/2026.
  let m = text.match(/^(\d{1,2})[-\/]\s*(\d{4})$/);
  if (m) {
    return ('0' + m[1]).slice(-2) + '-' + m[2];
  }

  return text;
}

function payrollRowToObject_(row) {
  const result = {};

  PAYROLL_HEADERS.forEach(function(header, index) {
    if (header === 'SalaryMonth') result[header] = normalizeSalaryMonth_(row[index]);
    else if (header !== 'ExtraData') result[header] = row[index];
  });
  if(row[PAYROLL_HEADERS.indexOf('ExtraData')]){ try { const extra=JSON.parse(String(row[PAYROLL_HEADERS.indexOf('ExtraData')])||'{}'); Object.keys(extra).forEach(k=>{ if(result[k]===undefined) result[k]=extra[k]; }); result.ExtraData=extra; } catch(e){ result.ExtraData={}; } }
  return result;
}

function createSession_(username, role) {
  const token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); const createdAt=new Date(); const expiresAt=new Date(createdAt.getTime()+CONFIG.SESSION_SECONDS*1000);
  const employeeCode=role==='EMPLOYEE'?username:''; const session={token:token,username:username,role:role,employeeCode:employeeCode,createdAt:createdAt.toISOString(),expiresAt:expiresAt.toISOString()};
  getSheet_(CONFIG.SHEETS.SESSIONS).appendRow([token,username,role,employeeCode,createdAt,expiresAt,true]);
  CacheService.getScriptCache().put('sess:'+token,JSON.stringify(session),Math.min(CONFIG.SESSION_SECONDS,CONFIG.SESSION_CACHE_SECONDS)); return token;
}
function requireEmployeeSession_(token){const s=findSession_(token);return s&&s.role==='EMPLOYEE'?s:null;}
function requireAdminSession_(token){const s=findSession_(token);return s&&s.role!=='EMPLOYEE'?s:null;}
function findSession_(token){
  if(!token)return null; const cache=CacheService.getScriptCache(), cached=cache.get('sess:'+token);
  if(cached){try{const s=JSON.parse(cached); if(new Date(s.expiresAt).getTime()>Date.now())return s;}catch(e){}}
  const sheet=getSheet_(CONFIG.SHEETS.SESSIONS), last=sheet.getLastRow(); if(last<=1)return null; const values=sheet.getRange(2,1,last-1,SESSION_HEADERS.length).getValues();
  for(let i=0;i<values.length;i++){ if(String(values[i][0])!==token)continue; if(!bool_(values[i][6]))return null; const expires=new Date(values[i][5]); if(isNaN(expires.getTime())||expires.getTime()<=Date.now()){sheet.getRange(i+2,7).setValue(false);return null;} const s={token:token,username:String(values[i][1]||''),role:String(values[i][2]||''),employeeCode:String(values[i][3]||''),createdAt:values[i][4],expiresAt:values[i][5]}; cache.put('sess:'+token,JSON.stringify({token:s.token,username:s.username,role:s.role,employeeCode:s.employeeCode,createdAt:new Date(s.createdAt).toISOString(),expiresAt:new Date(s.expiresAt).toISOString()}),Math.min(CONFIG.SESSION_SECONDS,CONFIG.SESSION_CACHE_SECONDS)); return s; } return null;
}

function findAdmin_(username) {
  const sheet=getSheet_(CONFIG.SHEETS.ADMINS), last=sheet.getLastRow(); if(last<=1)return null; const values=sheet.getRange(2,1,last-1,ADMIN_HEADERS.length).getValues();
  for(let i=0;i<values.length;i++){ if(String(values[i][0]||'').trim()===String(username).trim()) return {row:i+2,username:String(values[i][0]||''),passwordHash:String(values[i][1]||''),passwordSalt:String(values[i][2]||''),active:bool_(values[i][3]),updatedAt:values[i][4]||'',role:String(values[i][5]||CONFIG.ADMIN_DEFAULT_ROLE),failedAttempts:Number(values[i][6]||0),lockedUntil:values[i][7]||'',mustChangePassword:bool_(values[i][8]),lastLoginAt:values[i][9]||'',lastPasswordChange:values[i][10]||''}; } return null;
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
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }

  let text = String(value)
    .replace(/\u00A0/g, ' ')
    .trim();

  if (!text) return 0;

  // Defensive handling for values produced by older VBA versions.
  // Accept .5 / -.5 as well as comma-decimal forms.
  if (/^-?\.[0-9]+$/.test(text)) {
    text = text.replace(/^(-?)\./, '$10.');
  }

  // If both separators exist, the last one is treated as the decimal
  // separator and the other one as a thousands separator.
  const comma = text.lastIndexOf(',');
  const dot = text.lastIndexOf('.');

  if (comma >= 0 && dot >= 0) {
    if (comma > dot) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (comma >= 0) {
    // A single comma followed by 1-4 digits is normally a decimal value
    // in Vietnamese Excel exports; otherwise treat commas as grouping.
    const tail = text.substring(comma + 1);
    if (/^[0-9]{1,4}$/.test(tail)) {
      text = text.replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  }

  if (/^-?\d+\.$/.test(text)) {
    text = text.slice(0, -1);
  }

  const number = Number(text);
  return isFinite(number) ? number : 0;
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

  let content = String(e.postData.contents || '');

  if (!content) {
    return {};
  }

  // Remove UTF-8 BOM and surrounding whitespace.
  content = content.replace(/^\uFEFF/, '').trim();

  // Repair a legacy VBA number serialization bug such as :.5 or :-.5.
  // JSON requires a leading zero (0.5 / -0.5). The regex only targets a
  // numeric token immediately after a JSON delimiter, not quoted text.
  content = content.replace(/([:\[,]\s*)(-?)\.([0-9]+)/g, function(_, prefix, sign, digits) { return prefix + sign + '0.' + digits; });

  // If a proxy/client added text around the JSON, keep only the JSON object.
  const first = content.indexOf('{');
  const last = content.lastIndexOf('}');

  if (first > 0 || last >= 0) {
    if (first >= 0 && last > first) {
      content = content.substring(first, last + 1);
    }
  }

  try {
    const parsed = JSON.parse(content);

    if (!parsed || typeof parsed !== 'object') {
      throw new Error('JSON root must be an object');
    }

    return parsed;
  } catch (error) {
    const message = String(error && error.message || error);
    throw new Error(
      'Invalid JSON request: ' + message +
      '. Prefix=' + content.substring(0, 500)
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
