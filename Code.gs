/**
 * PAYROLL PORTAL - Google Apps Script
 * Version: Optimized batch sync
 *
 * Sheets:
 *   Users
 *   Salary
 *   SyncLog
 *
 * Script Properties:
 *   SPREADSHEET_ID (optional)
 *   SYNC_API_KEY   (required for VBA sync)
 *   ADMIN_PASSWORD_HASH / ADMIN_PASSWORD_SALT
 */

const CFG = {
  TZ: 'Asia/Ho_Chi_Minh',
  SESSION_TTL_SECONDS: 1800,
  SYNC_STATE_TTL_SECONDS: 3600,
  MIN_PASSWORD: 8,
  SHEETS: {
    USERS: 'Users',
    SALARY: 'Salary',
    SYNC_LOG: 'SyncLog'
  }
};

const USER_HEADERS = [
  'Username','EmployeeID','FullName','PasswordHash','PasswordSalt',
  'MustChangePassword','Active','UpdatedAt'
];

const SALARY_HEADERS = [
  'PayMonth','Factory','UpdatedAt','EmployeeID','FullName','CitizenID',
  'Department','Section','Position','TotalIncome','MonthlySalary','BasicSalary',
  'WorkingDays','HolidayDays','PaidLeaveDays','UnpaidLeaveDays',
  'RegionalMinimumLeaveDays','OvertimePay','OvertimeHours','RestDayHours',
  'OvertimeRestDayHours','NightRestDayOvertimeHours','HolidayHours',
  'HolidayOvertimeHours','NightHolidayOvertimeHours','NightShiftDays',
  'NightOvertimeHours','OtherMoney','Discipline','Loyalty2Years','Loyalty5Years',
  'Loyalty10Years','Housing','Transportation','AttendanceBonus',
  'SalesCommissionBonus','SeveranceUnusedLeave','SocialInsurance',
  'HealthInsurance','UnemploymentInsurance','PersonalIncomeTax','Advance',
  'OtherDeductions','NetPay'
];

const SYNC_HEADERS = [
  'UpdatedAt','Factory','PayMonth','BatchID','BatchNo','TotalBatches',
  'Rows','Inserted','Updated','Status','Message'
];

function doGet() {
  return json_({ok:true, service:'Payroll Portal', time:now_()});
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents) : {};

    switch (String(body.action || '').trim()) {
      case 'login': return json_(login_(body));
      case 'changePassword': return json_(changePassword_(body));
      case 'getPayslip': return json_(getPayslip_(body));
      case 'adminLogin': return json_(adminLogin_(body));
      case 'adminReset': return json_(adminReset_(body));
      case 'syncPayroll': return json_(syncPayroll_(body));
      case 'health': return json_({ok:true,time:now_()});
      default: return json_({ok:false,error:'INVALID_ACTION'});
    }
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({
      ok:false,
      error:'SERVER_ERROR',
      message:String(err && err.message ? err.message : err)
    });
  }
}

/* ================= AUTH ================= */

function login_(p) {
  const username = normalizeId_(p.username);
  const password = String(p.password || '');

  if (!username || !password)
    return {ok:false,error:'INVALID_CREDENTIALS'};

  const found = findUser_(username);
  if (!found || !toBool_(found.user.Active))
    return {ok:false,error:'INVALID_CREDENTIALS'};

  if (!verifyPassword_(password, found.user.PasswordSalt, found.user.PasswordHash))
    return {ok:false,error:'INVALID_CREDENTIALS'};

  const token = createSession_({
    type:'USER',
    username:username,
    row:found.row
  });

  return {
    ok:true,
    token:token,
    user:{
      username:username,
      employeeId:String(found.user.EmployeeID || ''),
      fullName:String(found.user.FullName || ''),
      mustChangePassword:toBool_(found.user.MustChangePassword)
    }
  };
}

function changePassword_(p) {
  const session = requireSession_(p.token,'USER');
  if (!session) return {ok:false,error:'SESSION_EXPIRED'};

  const p1 = String(p.newPassword || '');
  const p2 = String(p.newPassword2 || '');

  if (p1.length < CFG.MIN_PASSWORD)
    return {ok:false,error:'PASSWORD_TOO_SHORT'};

  if (p1 !== p2)
    return {ok:false,error:'PASSWORD_MISMATCH'};

  const sheet = getSheet_(CFG.SHEETS.USERS);
  ensureHeaders_(sheet,USER_HEADERS);

  const row = session.row;
  const values = sheet.getRange(row,1,1,USER_HEADERS.length).getValues()[0];

  if (normalizeId_(p1) === normalizeId_(values[1]))
    return {ok:false,error:'PASSWORD_CANNOT_BE_EMPLOYEE_ID'};

  const salt = Utilities.getUuid();
  values[3] = hashPassword_(p1,salt);
  values[4] = salt;
  values[5] = false;
  values[7] = now_();

  sheet.getRange(row,1,1,USER_HEADERS.length).setValues([values]);

  return {
    ok:true,
    user:{
      username:String(values[0] || ''),
      employeeId:String(values[1] || ''),
      fullName:String(values[2] || ''),
      mustChangePassword:false
    }
  };
}

/* ================= PAYSLIP ================= */

function getPayslip_(p) {
  const session = requireSession_(p.token,'USER');
  if (!session) return {ok:false,error:'SESSION_EXPIRED'};

  const found = findUser_(session.username);
  if (!found || !toBool_(found.user.Active))
    return {ok:false,error:'INVALID_SESSION_USER'};

  // Security check is server-side, not only frontend.
  if (toBool_(found.user.MustChangePassword))
    return {ok:false,error:'PASSWORD_CHANGE_REQUIRED'};

  const payMonth = normalizePayMonth_(p.payMonth);
  if (!payMonth) return {ok:false,error:'INVALID_MONTH'};

  const employeeId = normalizeId_(found.user.EmployeeID);
  const sheet = getSheet_(CFG.SHEETS.SALARY);
  ensureHeaders_(sheet,SALARY_HEADERS);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2)
    return {ok:false,error:'PAYSLIP_NOT_FOUND'};

  const data = sheet.getRange(
    2,1,lastRow-1,SALARY_HEADERS.length
  ).getValues();

  for (let i=0;i<data.length;i++) {
    const r=data[i];
    if (
      normalizePayMonth_(r[0])===payMonth &&
      normalizeId_(r[3])===employeeId
    ) {
      return {
        ok:true,
        payMonth:payMonth,
        employee:{
          employeeId:String(r[3] || ''),
          fullName:String(r[4] || ''),
          citizenId:String(r[5] || ''),
          department:String(r[6] || ''),
          section:String(r[7] || ''),
          position:String(r[8] || '')
        },
        payslip:rowToObject_(SALARY_HEADERS,r)
      };
    }
  }

  return {ok:false,error:'PAYSLIP_NOT_FOUND'};
}

/* ================= ADMIN ================= */

function setupSystem() {
  const ss = getSS_();
  if (!ss) throw new Error('Spreadsheet not found.');

  const users = getSheet_(CFG.SHEETS.USERS);
  const salary = getSheet_(CFG.SHEETS.SALARY);
  const syncLog = getSheet_(CFG.SHEETS.SYNC_LOG);

  ensureHeaders_(users,USER_HEADERS);
  ensureHeaders_(salary,SALARY_HEADERS);
  ensureHeaders_(syncLog,SYNC_HEADERS);

  // Keep identifiers as text so leading zeroes in CitizenID are preserved.
  users.getRange('A:B').setNumberFormat('@');
  salary.getRange('D:F').setNumberFormat('@');

  users.setFrozenRows(1);
  salary.setFrozenRows(1);
  syncLog.setFrozenRows(1);

  users.autoResizeColumns(1,USER_HEADERS.length);
  salary.autoResizeColumns(1,SALARY_HEADERS.length);
  syncLog.autoResizeColumns(1,SYNC_HEADERS.length);

  return {
    ok:true,
    spreadsheetId:ss.getId(),
    sheets:[CFG.SHEETS.USERS,CFG.SHEETS.SALARY,CFG.SHEETS.SYNC_LOG]
  };
}

function setAdminPassword(password) {
  password=String(password || '');
  if (password.length < CFG.MIN_PASSWORD)
    throw new Error('Admin password must be at least '+CFG.MIN_PASSWORD+' characters.');

  const salt=Utilities.getUuid();

  PropertiesService.getScriptProperties().setProperties({
    ADMIN_PASSWORD_HASH:hashPassword_(password,salt),
    ADMIN_PASSWORD_SALT:salt
  },true);

  return 'Admin password updated.';
}

function adminLogin_(p) {
  const password=String(p.password || '');
  const props=PropertiesService.getScriptProperties();

  const hash=props.getProperty('ADMIN_PASSWORD_HASH') || '';
  const salt=props.getProperty('ADMIN_PASSWORD_SALT') || '';

  if (!hash || !salt || !verifyPassword_(password,salt,hash))
    return {ok:false,error:'INVALID_ADMIN_CREDENTIALS'};

  return {
    ok:true,
    token:createSession_({type:'ADMIN'})
  };
}

function adminReset_(p) {
  const session=requireSession_(p.token,'ADMIN');
  if (!session) return {ok:false,error:'SESSION_EXPIRED'};

  const username=normalizeId_(p.username);
  const found=findUser_(username);

  if (!found || !toBool_(found.user.Active))
    return {ok:false,error:'USER_NOT_FOUND'};

  const employeeId=String(found.user.EmployeeID || '').trim();
  if (!employeeId) return {ok:false,error:'EMPLOYEE_ID_MISSING'};

  const salt=Utilities.getUuid();
  const sheet=getSheet_(CFG.SHEETS.USERS);
  const values=sheet.getRange(found.row,1,1,USER_HEADERS.length).getValues()[0];

  values[3]=hashPassword_(employeeId,salt);
  values[4]=salt;
  values[5]=true;
  values[7]=now_();

  sheet.getRange(found.row,1,1,USER_HEADERS.length).setValues([values]);

  return {
    ok:true,
    message:'PASSWORD_RESET',
    employeeId:employeeId,
    mustChangePassword:true
  };
}

/* ================= PAYROLL SYNC =================
 * Batch 1 clears the target factory/month ONCE.
 * Every batch then UPSERTS by Factory + PayMonth + EmployeeID.
 *
 * Therefore:
 *   - batch 1 cannot erase batch 2
 *   - retrying a failed batch is safe
 *   - rerunning a complete sync starts cleanly again at batch 1
 */

function syncPayroll_(p) {
  const expected = PropertiesService.getScriptProperties()
    .getProperty('SYNC_API_KEY') || '';

  if (!expected || String(p.apiKey || '') !== expected)
    return {ok:false,error:'INVALID_API_KEY'};

  const factory = String(p.factory || '').trim();
  const payMonth = normalizePayMonth_(p.payMonth);
  const rows = Array.isArray(p.rows) ? p.rows : [];
  const batchId = String(p.batchId || '').trim();
  const batchNo = Number(p.batchNo);
  const totalBatches = Number(p.totalBatches);
  const isFirstBatch = p.isFirstBatch === true;

  if (!['Snack','Flexible'].includes(factory))
    return {ok:false,error:'INVALID_FACTORY'};

  if (!payMonth)
    return {ok:false,error:'INVALID_MONTH'};

  if (!rows.length)
    return {ok:false,error:'EMPTY_BATCH'};

  if (!batchId)
    return {ok:false,error:'BATCH_ID_REQUIRED'};

  if (!Number.isInteger(batchNo) || batchNo < 1)
    return {ok:false,error:'INVALID_BATCH_NO'};

  if (!Number.isInteger(totalBatches) ||
      totalBatches < 1 ||
      batchNo > totalBatches)
    return {ok:false,error:'INVALID_TOTAL_BATCHES'};

  if (batchNo === 1 && !isFirstBatch)
    return {ok:false,error:'FIRST_BATCH_REQUIRED'};

  if (batchNo > 1 && isFirstBatch)
    return {ok:false,error:'INVALID_FIRST_BATCH_FLAG'};

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_(CFG.SHEETS.SALARY);
    ensureHeaders_(sheet,SALARY_HEADERS);

    const stateKey = syncStateKey_(factory,payMonth);
    const props = PropertiesService.getScriptProperties();
    const state = readSyncState_(props,stateKey);

    // A sync is treated as a small state machine:
    // 1) batch 1 clears the period once;
    // 2) later batches must arrive in order;
    // 3) retrying an already accepted batch is idempotent;
    // 4) another batchId cannot overwrite an active sync.
    if (batchNo === 1) {
      if (state && state.batchId !== batchId) {
        return {
          ok:false,
          error:'SYNC_IN_PROGRESS',
          batchNo:state.lastBatchNo || 0,
          totalBatches:state.totalBatches || totalBatches
        };
      }

      if (!state || state.batchId !== batchId) {
        clearPeriod_(sheet,factory,payMonth);
        writeSyncState_(props,stateKey,{
          batchId:batchId,
          factory:factory,
          payMonth:payMonth,
          lastBatchNo:0,
          totalBatches:totalBatches
        });
      }
    } else {
      if (!state || state.batchId !== batchId)
        return {ok:false,error:'SYNC_STATE_MISSING'};

      if (state.totalBatches !== totalBatches)
        return {ok:false,error:'TOTAL_BATCHES_MISMATCH'};

      if (batchNo > Number(state.lastBatchNo || 0) + 1)
        return {
          ok:false,
          error:'BATCH_OUT_OF_ORDER',
          expectedBatch:Number(state.lastBatchNo || 0) + 1
        };
    }

    const result = upsertPayrollRows_(sheet,factory,payMonth,rows);

    const currentLast = state && state.batchId === batchId
      ? Number(state.lastBatchNo || 0)
      : 0;

    if (batchNo > currentLast) {
      writeSyncState_(props,stateKey,{
        batchId:batchId,
        factory:factory,
        payMonth:payMonth,
        lastBatchNo:batchNo,
        totalBatches:totalBatches
      });
    }

    if (batchNo === totalBatches) {
      props.deleteProperty(stateKey);
    }

    logSync_({
      factory:factory,
      payMonth:payMonth,
      batchId:batchId,
      batchNo:batchNo,
      totalBatches:totalBatches,
      rows:rows.length,
      inserted:result.inserted,
      updated:result.updated,
      status:'OK',
      message:batchNo === totalBatches ? 'SYNC_COMPLETE' : 'BATCH_OK'
    });

    return {
      ok:true,
      factory:factory,
      payMonth:payMonth,
      batchId:batchId,
      batchNo:batchNo,
      totalBatches:totalBatches,
      inserted:result.inserted,
      updated:result.updated,
      received:rows.length,
      complete:batchNo === totalBatches
    };
  } catch(err) {
    logSync_({
      factory:factory,
      payMonth:payMonth,
      batchId:batchId,
      batchNo:batchNo,
      totalBatches:totalBatches,
      rows:rows.length,
      inserted:0,
      updated:0,
      status:'ERROR',
      message:String(err && err.message ? err.message : err)
    });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function syncStateKey_(factory,payMonth) {
  return 'SYNC_STATE_' +
    String(factory).trim().toUpperCase() + '_' +
    normalizePayMonth_(payMonth).replace('-','_');
}

function readSyncState_(props,key) {
  const raw = props.getProperty(key);
  if (!raw) return null;

  try {
    const state = JSON.parse(raw);
    if (!state || !state.batchId) return null;

    const age = Date.now() - Number(state.updatedAt || 0);
    if (!Number.isFinite(age) || age < 0 ||
        age > CFG.SYNC_STATE_TTL_SECONDS * 1000) {
      props.deleteProperty(key);
      return null;
    }

    return state;
  } catch (_) {
    props.deleteProperty(key);
    return null;
  }
}

function writeSyncState_(props,key,state) {
  state.updatedAt = Date.now();
  props.setProperty(key,JSON.stringify(state));
}

function upsertPayrollRows_(sheet,factory,payMonth,rows) {
  const lastRow=sheet.getLastRow();

  const existing=lastRow>=2
    ? sheet.getRange(2,1,lastRow-1,SALARY_HEADERS.length).getValues()
    : [];

  const keyToRow=new Map();

  for(let i=0;i<existing.length;i++){
    const r=existing[i];
    const employeeId=normalizeId_(r[3]);
    if(!employeeId) continue;

    const key=payrollKey_(r[1],r[0],employeeId);
    if(!keyToRow.has(key))
      keyToRow.set(key,i+2);
  }

  const updates=[];
  const inserts=[];
  const seen=new Set();

  rows.forEach(item=>{
    const row=normalizePayrollRow_(item,factory,payMonth);
    const key=payrollKey_(factory,payMonth,row[3]);

    if(!row[3] || seen.has(key)) return;
    seen.add(key);

    const existingRow=keyToRow.get(key);

    if(existingRow)
      updates.push({row:existingRow,values:row});
    else
      inserts.push(row);
  });

  // Write existing rows in contiguous blocks instead of one setValue per row.
  updates.sort((a,b)=>a.row-b.row);

  let start=null;
  let values=[];
  let previous=null;

  const flush=()=>{
    if(start===null) return;
    sheet.getRange(
      start,1,values.length,SALARY_HEADERS.length
    ).setValues(values);

    start=null;
    values=[];
    previous=null;
  };

  updates.forEach(u=>{
    if(start===null){
      start=u.row;
      values=[u.values];
    } else if(u.row===previous+1){
      values.push(u.values);
    } else {
      flush();
      start=u.row;
      values=[u.values];
    }
    previous=u.row;
  });

  flush();

  if(inserts.length){
    const s=sheet.getLastRow()+1;
    sheet.getRange(s,4,inserts.length,3).setNumberFormat('@');
    sheet.getRange(
      s,1,inserts.length,SALARY_HEADERS.length
    ).setValues(inserts);

    // CitizenID must stay text to preserve leading zero.
    sheet.getRange(s,6,inserts.length,1).setNumberFormat('@');
  }

  return {
    inserted:inserts.length,
    updated:updates.length
  };
}

function normalizePayrollRow_(item,factory,payMonth) {
  const n=v=>{
    if(v===null || v===undefined || v==='') return 0;
    const x=Number(v);
    return Number.isFinite(x)?x:0;
  };

  const s=v=>{
    if(v===null || v===undefined) return '';
    return String(v).trim().replace(/^'+/,'');
  };

  return [
    payMonth,
    factory,
    now_(),
    s(item.EmployeeID),
    s(item.FullName),
    s(item.CitizenID),
    s(item.Department),
    s(item.Section),
    s(item.Position),
    n(item.TotalIncome),
    n(item.MonthlySalary),
    n(item.BasicSalary),
    n(item.WorkingDays),
    n(item.HolidayDays),
    n(item.PaidLeaveDays),
    n(item.UnpaidLeaveDays),
    n(item.RegionalMinimumLeaveDays),
    n(item.OvertimePay),
    n(item.OvertimeHours),
    n(item.RestDayHours),
    n(item.OvertimeRestDayHours),
    n(item.NightRestDayOvertimeHours),
    n(item.HolidayHours),
    n(item.HolidayOvertimeHours),
    n(item.NightHolidayOvertimeHours),
    n(item.NightShiftDays),
    n(item.NightOvertimeHours),
    n(item.OtherMoney),
    n(item.Discipline),
    n(item.Loyalty2Years),
    n(item.Loyalty5Years),
    n(item.Loyalty10Years),
    n(item.Housing),
    n(item.Transportation),
    n(item.AttendanceBonus),
    n(item.SalesCommissionBonus),
    n(item.SeveranceUnusedLeave),
    n(item.SocialInsurance),
    n(item.HealthInsurance),
    n(item.UnemploymentInsurance),
    n(item.PersonalIncomeTax),
    n(item.Advance),
    n(item.OtherDeductions),
    n(item.NetPay)
  ];
}

function clearPeriod_(sheet,factory,payMonth) {
  const last = sheet.getLastRow();
  if (last < 2) return;

  const data = sheet.getRange(2,1,last-1,2).getValues();
  const rows = [];

  for (let i=0;i<data.length;i++) {
    if (
      String(data[i][1] || '').trim() === factory &&
      normalizePayMonth_(data[i][0]) === payMonth
    ) {
      rows.push(i + 2);
    }
  }

  if (!rows.length) return;

  // Delete from bottom to top. This avoids leaving thousands of blank
  // rows behind after every monthly resync.
  let start = rows[rows.length - 1];
  let count = 1;

  for (let i=rows.length - 2;i>=0;i--) {
    if (rows[i] === start - 1) {
      start = rows[i];
      count++;
    } else {
      sheet.deleteRows(start,count);
      start = rows[i];
      count = 1;
    }
  }

  sheet.deleteRows(start,count);
}

function payrollKey_(factory,payMonth,employeeId) {
  return String(factory||'').trim().toUpperCase()+'|'+
    normalizePayMonth_(payMonth)+'|'+
    normalizeId_(employeeId);
}

/* ================= USERS ================= */

function syncUsers_(employees) {
  const sheet=getSheet_(CFG.SHEETS.USERS);
  ensureHeaders_(sheet,USER_HEADERS);
  sheet.getRange('A:B').setNumberFormat('@');

  if(!Array.isArray(employees) || !employees.length)
    return {inserted:0,updated:0};

  const last=sheet.getLastRow();
  const data=last>=2
    ? sheet.getRange(2,1,last-1,USER_HEADERS.length).getValues()
    : [];

  const map=new Map();

  data.forEach((r,i)=>{
    const username=normalizeId_(r[0]);
    if(username) map.set(username,i+2);
  });

  let inserted=0;
  let updated=0;

  employees.forEach(emp=>{
    const username=normalizeId_(emp.CitizenID || emp.Username);
    const employeeId=String(emp.EmployeeID || '').trim();
    const fullName=String(emp.FullName || '').trim();
    const active=emp.Active===undefined ? true : !!emp.Active;

    if(!username || !employeeId) return;

    const row=map.get(username);

    if(row){
      const v=sheet.getRange(
        row,1,1,USER_HEADERS.length
      ).getValues()[0];

      v[1]=employeeId;
      v[2]=fullName;
      v[6]=active;
      v[7]=now_();

      sheet.getRange(
        row,1,1,USER_HEADERS.length
      ).setValues([v]);

      updated++;
    } else {
      const salt=Utilities.getUuid();

      const v=[
        username,
        employeeId,
        fullName,
        hashPassword_(employeeId,salt),
        salt,
        true,
        active,
        now_()
      ];

      sheet.getRange(
        sheet.getLastRow()+1,
        1,1,USER_HEADERS.length
      ).setValues([v]);

      inserted++;
    }
  });

  return {inserted:inserted,updated:updated};
}

/* ================= SESSION ================= */

function createSession_(data) {
  const token=Utilities.getUuid()+'-'+Utilities.getUuid();

  CacheService.getScriptCache().put(
    'SESSION_'+token,
    JSON.stringify(data),
    CFG.SESSION_TTL_SECONDS
  );

  return token;
}

function requireSession_(token,type) {
  if(!token) return null;

  const key='SESSION_'+String(token);
  const raw=CacheService.getScriptCache().get(key);

  if(!raw) return null;

  try {
    const data=JSON.parse(raw);

    if(data.type!==type) return null;

    // Sliding TTL.
    CacheService.getScriptCache().put(
      key,JSON.stringify(data),CFG.SESSION_TTL_SECONDS
    );

    return data;
  } catch(e) {
    return null;
  }
}

/* ================= SHEETS ================= */

function findUser_(username) {
  const sheet=getSheet_(CFG.SHEETS.USERS);
  ensureHeaders_(sheet,USER_HEADERS);

  const last=sheet.getLastRow();
  if(last<2) return null;

  const data=sheet.getRange(
    2,1,last-1,USER_HEADERS.length
  ).getValues();

  const target=normalizeId_(username);

  for(let i=0;i<data.length;i++){
    if(normalizeId_(data[i][0])===target){
      const user={};
      USER_HEADERS.forEach((h,j)=>user[h]=data[i][j]);
      return {row:i+2,user:user};
    }
  }

  return null;
}

function getSS_() {
  const id=PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  return id
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const ss=getSS_();
  if(!ss) throw new Error('Spreadsheet not found.');

  let sheet=ss.getSheetByName(name);
  if(!sheet) sheet=ss.insertSheet(name);

  return sheet;
}

function ensureHeaders_(sheet,headers) {
  if(sheet.getMaxColumns()<headers.length){
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length-sheet.getMaxColumns()
    );
  }

  const current=sheet.getRange(
    1,1,1,headers.length
  ).getValues()[0];

  const same=headers.every(
    (h,i)=>String(current[i]||'').trim()===h
  );

  if(!same)
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
}

function logSync_(x) {
  const sheet=getSheet_(CFG.SHEETS.SYNC_LOG);
  ensureHeaders_(sheet,SYNC_HEADERS);

  sheet.appendRow([
    now_(),
    x.factory||'',
    x.payMonth||'',
    x.batchId||'',
    x.batchNo||'',
    x.totalBatches||'',
    x.rows||0,
    x.inserted||0,
    x.updated||0,
    x.status||'',
    x.message||''
  ]);
}

/* ================= HELPERS ================= */

function normalizeId_(v) {
  if(v===null || v===undefined) return '';
  return String(v).trim().replace(/^'+/,'');
}

function normalizePayMonth_(v) {
  const s=String(v||'').trim();

  if(!/^\d{2}-\d{4}$/.test(s))
    return '';

  const mm=Number(s.slice(0,2));

  return mm>=1 && mm<=12 ? s : '';
}

function toBool_(v) {
  if(v===true) return true;

  return [
    'true','1','yes','y'
  ].includes(
    String(v||'').toLowerCase().trim()
  );
}

function rowToObject_(headers,row) {
  const o={};
  headers.forEach((h,i)=>o[h]=row[i]);
  return o;
}

function hashPassword_(password,salt) {
  const raw=Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt)+':'+String(password),
    Utilities.Charset.UTF_8
  );

  return raw.map(b=>{
    const n=b<0?b+256:b;
    return ('0'+n.toString(16)).slice(-2);
  }).join('');
}

function verifyPassword_(password,salt,hash) {
  if(!salt || !hash) return false;
  return hashPassword_(password,salt)===String(hash);
}

function now_() {
  return Utilities.formatDate(
    new Date(),
    CFG.TZ,
    'dd/MM/yyyy HH:mm:ss'
  );
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run once manually after pasting Code.gs.
 * Example:
 *   setAdminPassword('your-admin-password');
 */
