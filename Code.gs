/**
 * GOOGLE APPS SCRIPT - SALARY API & DATABASE CONTROLLER
 * URL: https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec
 */

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    
    if (action === 'syncData') {
      return handleSyncData(contents);
    }
    
    return responseJSON({ success: false, message: 'Invalid Action' });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const ma_nv = e.parameter.ma_nv;
    const cccd = e.parameter.cccd;
    const period = e.parameter.period; // Ví dụ: '08-2026'

    if (action === 'getSalary') {
      return handleGetSalary(ma_nv, cccd, period);
    }
    
    if (action === 'getPeriods') {
      return handleGetPeriods();
    }

    return responseJSON({ success: false, message: 'Invalid Endpoint' });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

// ---------------------------------------------------------
// HELPER FUNCTIONS
// ---------------------------------------------------------

function handleSyncData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const period = payload.period; // e.g., '08-2026'
  const data = payload.data;
  
  let sheet = ss.getSheetByName(period);
  if (!sheet) {
    sheet = ss.insertSheet(period);
  } else {
    sheet.clearContents(); // Ghi đè dữ liệu kỳ đó nếu đã tồn tại
  }

  if (data.length === 0) {
    return responseJSON({ success: true, message: 'No data synced' });
  }

  // Format CCCD dạng Text
  const keys = Object.keys(data[0]);
  const rows = [keys]; // Header

  data.forEach(item => {
    const row = keys.map(k => {
      let val = item[k];
      if (k === 'cccd') return "'" + String(val); // Bảo toàn số 0 ở đầu cho CCCD
      return val;
    });
    rows.push(row);
  });

  sheet.getRange(1, 1, rows.length, keys.length).setValues(rows);
  // Định dạng cột CCCD là Plain Text
  const cccdColIdx = keys.indexOf('cccd') + 1;
  if (cccdColIdx > 0) {
    sheet.getRange(2, cccdColIdx, rows.length - 1, 1).setNumberFormat('@');
  }

  return responseJSON({ success: true, message: `Đã đồng bộ thành công ${data.length} dòng dữ liệu kỳ ${period}` });
}

function handleGetSalary(ma_nv, cccd, period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Nếu không chỉ định period, lấy sheet mới nhất
  let sheet = period ? ss.getSheetByName(period) : null;
  if (!sheet) {
    const sheets = ss.getSheets();
    sheet = sheets[0]; // Mặc định lấy sheet đầu tiên
  }

  if (!sheet) {
    return responseJSON({ success: false, message: 'Chưa có dữ liệu bảng lương!' });
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return responseJSON({ success: false, message: 'Dữ liệu trống!' });
  }

  const headers = values[0];
  const maNvIdx = headers.indexOf('ma_nv');
  const cccdIdx = headers.indexOf('cccd');

  const cleanInputMaNV = String(ma_nv || '').trim().toLowerCase();
  const cleanInputCCCD = String(cccd || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowMaNV = String(row[maNvIdx] || '').trim().toLowerCase();
    let rowCCCD = String(row[cccdIdx] || '').trim().toLowerCase();
    if (rowCCCD.startsWith("'")) rowCCCD = rowCCCD.substring(1);

    if (rowMaNV === cleanInputMaNV && rowCCCD === cleanInputCCCD) {
      let result = {};
      headers.forEach((h, idx) => {
        result[h] = row[idx];
      });
      result.period = sheet.getName();
      return responseJSON({ success: true, data: result });
    }
  }

  return responseJSON({ success: false, message: 'Thông tin Mã Nhân Viên hoặc Số CCCD không chính xác!' });
}

function handleGetPeriods() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const periods = sheets.map(s => s.getName()).filter(name => name.includes('-'));
  return responseJSON({ success: true, periods: periods });
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
