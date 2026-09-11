const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const API_KEY = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE"; // Khớp với VBA

const SHEET_USERS = "USERS";
const SHEET_PAYROLL = "PAYROLL";

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;

    if (action === "syncPayroll") {
      return handleSyncPayroll(contents);
    } else if (action === "login") {
      return handleLogin(contents);
    } else if (action === "changePassword") {
      return handleChangePassword(contents);
    } else if (action === "adminResetPassword") {
      return handleAdminResetPassword(contents);
    }

    return jsonResponse({ ok: false, message: "Hành động không hợp lệ" });
  } catch (err) {
    return jsonResponse({ ok: false, message: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, message: "Payroll API System Running" });
}

// ------------------------------------------------------------
// HANDLERS
// ------------------------------------------------------------

function handleSyncPayroll(data) {
  if (data.apiKey !== API_KEY) {
    return jsonResponse({ ok: false, message: "Khóa API không hợp lệ" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetUsers = getOrCreateSheet(ss, SHEET_USERS, ["CitizenID", "EmployeeID", "PasswordHash", "MustChangePassword", "Role", "UpdatedAt"]);
  const sheetPayroll = getOrCreateSheet(ss, SHEET_PAYROLL, [
    "PayMonth", "Factory", "CitizenID", "EmployeeID", "FullName", "Department", "Section", "Position",
    "TotalIncome", "MonthlySalary", "BasicSalary", "WorkingDays", "HolidayDays", "PaidLeaveDays",
    "UnpaidLeaveDays", "RegionalMinimumLeaveDays", "OvertimePay", "OvertimeHours", "RestDayHours",
    "OvertimeRestDayHours", "NightRestDayOvertimeHours", "HolidayHours", "HolidayOvertimeHours",
    "NightHolidayOvertimeHours", "NightShiftDays", "NightOvertimeHours", "OtherMoney", "Discipline",
    "Loyalty2Years", "Loyalty5Years", "Loyalty10Years", "Housing", "Transportation", "AttendanceBonus",
    "SalesCommissionBonus", "SeveranceUnusedLeave", "SocialInsurance", "HealthInsurance",
    "UnemploymentInsurance", "PersonalIncomeTax", "Advance", "OtherDeductions", "NetPay", "UpdatedAt"
  ]);

  const nowStr = formatDate(new Date());

  // Đồng bộ Users
  const userMap = getUserMap(sheetUsers);
  const newUsers = [];
  
  // Xóa dữ liệu cũ nếu là lô đầu
  if (data.isFirstBatch) {
    clearPayrollData(sheetPayroll, data.payMonth, data.factory);
  }

  const payrollRows = [];

  data.rows.forEach(r => {
    // 1. Thêm user mới nếu chưa tồn tại
    if (r.CitizenID && !userMap.has(r.CitizenID)) {
      const defaultHash = hashPassword(r.EmployeeID);
      newUsers.push([r.CitizenID, r.EmployeeID, defaultHash, true, "USER", nowStr]);
      userMap.set(r.CitizenID, true);
    }

    // 2. Map dữ liệu lương
    payrollRows.push([
      data.payMonth, data.factory, "'" + r.CitizenID, r.EmployeeID, r.FullName, r.Department, r.Section, r.Position,
      r.TotalIncome, r.MonthlySalary, r.BasicSalary, r.WorkingDays, r.HolidayDays, r.PaidLeaveDays,
      r.UnpaidLeaveDays, r.RegionalMinimumLeaveDays, r.OvertimePay, r.OvertimeHours, r.RestDayHours,
      r.OvertimeRestDayHours, r.NightRestDayOvertimeHours, r.HolidayHours, r.HolidayOvertimeHours,
      r.NightHolidayOvertimeHours, r.NightShiftDays, r.NightOvertimeHours, r.OtherMoney, r.Discipline,
      r.Loyalty2Years, r.Loyalty5Years, r.Loyalty10Years, r.Housing, r.Transportation, r.AttendanceBonus,
      r.SalesCommissionBonus, r.SeveranceUnusedLeave, r.SocialInsurance, r.HealthInsurance,
      r.UnemploymentInsurance, r.PersonalIncomeTax, r.Advance, r.OtherDeductions, r.NetPay, nowStr
    ]);
  });

  if (newUsers.length > 0) {
    sheetUsers.getRange(sheetUsers.getLastRow() + 1, 1, newUsers.length, newUsers[0].length).setValues(newUsers);
  }

  if (payrollRows.length > 0) {
    sheetPayroll.getRange(sheetPayroll.getLastRow() + 1, 1, payrollRows.length, payrollRows[0].length).setValues(payrollRows);
  }

  return jsonResponse({ ok: true, message: `Lô ${data.batchNo}/${data.totalBatches} đã xử lý thành công.` });
}

function handleLogin(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  const uData = sheetUsers.getDataRange().getValues();

  const cccd = String(data.citizenId).trim();
  const pwd = String(data.password).trim();
  const inputHash = hashPassword(pwd);

  for (let i = 1; i < uData.length; i++) {
    const rowCCCD = String(uData[i][0]).replace("'", "").trim();
    if (rowCCCD === cccd) {
      if (uData[i][2] === inputHash) {
        const mustChange = uData[i][3] === true || uData[i][3] === "TRUE";
        const role = uData[i][4] || "USER";
        
        let payrolls = [];
        if (!mustChange) {
          payrolls = getUserPayrolls(ss, cccd);
        }

        return jsonResponse({
          ok: true,
          mustChangePassword: mustChange,
          role: role,
          employeeId: uData[i][1],
          citizenId: cccd,
          payrolls: payrolls
        });
      } else {
        return jsonResponse({ ok: false, message: "Mật khẩu không chính xác." });
      }
    }
  }

  return jsonResponse({ ok: false, message: "Số Căn cước công dân không tồn tại." });
}

function handleChangePassword(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  const uData = sheetUsers.getDataRange().getValues();

  const cccd = String(data.citizenId).trim();
  const oldHash = hashPassword(String(data.oldPassword).trim());
  const newHash = hashPassword(String(data.newPassword).trim());

  for (let i = 1; i < uData.length; i++) {
    const rowCCCD = String(uData[i][0]).replace("'", "").trim();
    if (rowCCCD === cccd) {
      if (uData[i][2] === oldHash) {
        const nowStr = formatDate(new Date());
        sheetUsers.getRange(i + 1, 3).setValue(newHash);
        sheetUsers.getRange(i + 1, 4).setValue(false);
        sheetUsers.getRange(i + 1, 6).setValue(nowStr);

        const payrolls = getUserPayrolls(ss, cccd);
        return jsonResponse({ ok: true, message: "Đổi mật khẩu thành công!", payrolls: payrolls });
      } else {
        return jsonResponse({ ok: false, message: "Mật khẩu cũ không chính xác." });
      }
    }
  }
  return jsonResponse({ ok: false, message: "Không tìm thấy người dùng." });
}

function handleAdminResetPassword(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  const uData = sheetUsers.getDataRange().getValues();

  const adminCCCD = String(data.adminCitizenId).trim();
  const targetCCCD = String(data.targetCitizenId).trim();

  // Kiểm tra quyền Admin
  let isAdmin = false;
  for (let i = 1; i < uData.length; i++) {
    if (String(uData[i][0]).replace("'", "").trim() === adminCCCD && uData[i][4] === "ADMIN") {
      isAdmin = true;
      break;
    }
  }

  if (!isAdmin) {
    return jsonResponse({ ok: false, message: "Bạn không có quyền Admin để thực hiện thao tác này." });
  }

  for (let i = 1; i < uData.length; i++) {
    const rowCCCD = String(uData[i][0]).replace("'", "").trim();
    if (rowCCCD === targetCCCD) {
      const empId = uData[i][1];
      const defaultHash = hashPassword(empId);
      const nowStr = formatDate(new Date());

      sheetUsers.getRange(i + 1, 3).setValue(defaultHash);
      sheetUsers.getRange(i + 1, 4).setValue(true);
      sheetUsers.getRange(i + 1, 6).setValue(nowStr);

      return jsonResponse({ ok: true, message: `Đã reset mật khẩu cho CCCD ${targetCCCD} về Mã NV (${empId}).` });
    }
  }

  return jsonResponse({ ok: false, message: "Không tìm thấy nhân viên cần reset." });
}

// ------------------------------------------------------------
// UTILS
// ------------------------------------------------------------

function getUserPayrolls(ss, cccd) {
  const sheetPayroll = ss.getSheetByName(SHEET_PAYROLL);
  if (!sheetPayroll) return [];
  
  const data = sheetPayroll.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const rowCCCD = String(data[i][2]).replace("'", "").trim();
    if (rowCCCD === cccd) {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = data[i][idx];
      });
      list.push(obj);
    }
  }

  // Sắp xếp tháng mới nhất lên đầu
  return list.reverse();
}

function hashPassword(str) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function formatDate(d) {
  return Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy HH:mm:ss");
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function getUserMap(sheetUsers) {
  const map = new Map();
  const data = sheetUsers.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    map.set(String(data[i][0]).replace("'", "").trim(), true);
  }
  return map;
}

function clearPayrollData(sheetPayroll, payMonth, factory) {
  const data = sheetPayroll.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === payMonth && String(data[i][1]) === factory) {
      sheetPayroll.deleteRow(i + 1);
    }
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
