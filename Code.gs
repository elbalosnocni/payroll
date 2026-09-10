/**
 * PAYROLL PORTAL - Google Apps Script (Optimized & Fixed)
 */

const CFG = {
  TZ: 'Asia/Ho_Chi_Minh',
  SESSION_TTL_SECONDS: 1800,
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

/* ================= PAYROLL SYNC ================= */

function syncPayroll_(p) {
  const expected=PropertiesService.getScriptProperties()
    .getProperty('SYNC_API_KEY') || '';

  if (!expected || String(p.apiKey || '') !== expected)
    return {ok:false,error:'INVALID_API_KEY'};

  const factory=String(p.factory || '').trim();
  const payMonth=normalizePayMonth_(p.payMonth);
  const rows=Array.isArray(p.rows) ? p.rows : [];
  const batchId=String(p.batchId || '').trim();
  const batchNo=Number(p.batchNo || 1);
  const totalBatches=Number(p.totalBatches || 1);
  const isFirstBatch=!!p.isFirstBatch;

  if (!['Snack','Flexible'].includes(factory))
    return {ok:false,error:'INVALID_FACTORY'};

  if (!payMonth)
    return {ok:false,error:'INVALID_MONTH'};

  if (!rows.length)
    return {ok:false,error:'EMPTY_BATCH'};

  if (!batchId)
    return {ok:false,error:'BATCH_ID_REQUIRED'};

  if (!Number.isInteger(batchNo) || batchNo<1)
    return {ok:false,error:'INVALID_BATCH_NO'};

  if (!Number.isInteger(totalBatches) || totalBatches<batchNo)
    return {ok:false,error:'INVALID_TOTAL_BATCHES'};

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet=getSheet_(CFG.SHEETS.SALARY);
    ensureHeaders_(sheet,SALARY_HEADERS);

    if (isFirstBatch)
      clearPeriod_(sheet,factory,payMonth);

    const result=upsertPayrollRows_(sheet,factory,payMonth,rows);

    // FIX: Tự động trích xuất nhân viên và đồng bộ sang bảng Users
    syncUsersFromPayrollRows_(rows);

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
      message:isFirstBatch
        ? 'FIRST_BATCH_CLEARED_PERIOD'
        : 'BATCH_OK'
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
      complete:batchNo===totalBatches
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

function syncUsersFromPayrollRows_(rows) {
  const userList = rows.map(r => ({
    CitizenID: r.CitizenID,
    EmployeeID: r.EmployeeID,
    FullName: r.FullName,
    Active: true
  })).filter(u => u.CitizenID && u.EmployeeID);

  if (userList.length > 0) {
    syncUsers_(userList);
  }
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
    sheet.getRange(
      s,1,inserts.length,SALARY_HEADERS.length
    ).setValues(inserts);

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
  const last=sheet.getLastRow();
  if(last<2) return;

  const data=sheet.getRange(2,1,last-1,2).getValues();
  const rows=[];

  for(let i=0;i<data.length;i++){
    if(
      String(data[i][1]||'').trim()===factory &&
      normalizePayMonth_(data[i][0])===payMonth
    ){
      rows.push(i+2);
    }
  }

  if(!rows.length) return;

  let start=rows[0];
  let count=1;

  for(let i=1;i<rows.length;i++){
    if(rows[i]===rows[i-1]+1){
      count++;
    } else {
      sheet.getRange(
        start,1,count,SALARY_HEADERS.length
      ).clearContent();

      start=rows[i];
      count=1;
    }
  }

  sheet.getRange(
    start,1,count,SALARY_HEADERS.length
  ).clearContent();
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
