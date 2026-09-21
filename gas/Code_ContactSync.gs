function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Không nhận được postData"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var jsonString = e.postData.contents;
    var payload;
    
    // Bắt lỗi cú pháp JSON tại đây
    try {
      payload = JSON.parse(jsonString);
    } catch (parseError) {
      return ContentService.createTextOutput(JSON.stringify({
        "status": "error", 
        "message": "Lỗi cú pháp JSON từ VBA: " + parseError.message,
        "raw_data_snippet": jsonString.substring(0, 300) // Trả lại 300 ký tự đầu để soi lỗi
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var targetData = payload.data; 
    
    if (targetData && targetData.length > 0) {
      // 1. Tạo Tiêu đề ở Dòng 1
      var header = [
        "Mã NV", "Họ và Tên", "Số CCCD", "Số điện thoại", "Email", 
        "Nơi công tác", "Bộ phận", "Chức vụ", "Thường trú", "Tạm trú", 
        "Ngày sinh", "Nhãn / Nhóm", "Trạng thái"
      ];
      sheet.getRange(1, 1, 1, 13).setValues([header]);
      
      // 2. Ép ghi dữ liệu từ Dòng 2
      var numRows = targetData.length;
      var numCols = targetData[0].length;
      
      sheet.getRange(2, 1, numRows, numCols).setValues(targetData);
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success", 
        "message": "Đã ghi thành công " + numRows + " dòng vào Sheet!"
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Mảng data rỗng"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.message}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function dongBoDanhBaTongHop() {
  var startTime = new Date().getTime();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Thông báo', 'Không có dữ liệu nhân viên để đồng bộ!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Đọc toàn bộ vùng dữ liệu từ Cột A đến Cột M (13 cột)
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 13);
  var data = dataRange.getValues();
  var displayValues = dataRange.getDisplayValues();
  
  var countCreated = 0;
  var countUpdated = 0;
  var labelCache = {};
  
  for (var i = 0; i < data.length; i++) {
    // Tránh vượt quá 5 phút execution time limit
    if (new Date().getTime() - startTime > 270000) { 
      SpreadsheetApp.getUi().alert('Tạm dừng', 'Đã chạy quá 4.5 phút! Hãy bấm Đồng bộ lại để tiếp tục phần còn lại.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var row = data[i];
    var displayRow = displayValues[i];
    
    // SỬA LỖI: Đọc chính xác giá trị theo từng cột tương ứng
    var maNV        = displayRow[0] ? displayRow[0].toString().trim() : ""; // Cột A
    var hoTen       = row[1] ? row[1].toString().trim() : "";              // Cột B
    var soCCCD      = displayRow[2] ? displayRow[2].toString().trim() : ""; // Cột C
    var soDienThoai = displayRow[3] ? displayRow[3].toString().trim() : ""; // Cột D
    var email       = row[4] ? row[4].toString().trim() : "";              // Cột E
    var congTy      = row[5] ? row[5].toString().trim() : "";              // Cột F
    var boPhan      = row[6] ? row[6].toString().trim() : "";              // Cột G
    var chucVu      = row[7] ? row[7].toString().trim() : "";              // Cột H
    var thuongTru   = row[8] ? row[8].toString().trim() : "";              // Cột I
    var tamTru      = row[9] ? row[9].toString().trim() : "";              // Cột J
    var ngaySinh    = row[10];                                             // Cột K
    var labelName   = row[11] ? row[11].toString().trim() : "";            // Cột L
    var trangThai   = displayRow[12] ? displayRow[12].toString().trim() : ""; // Cột M
    
    if (!hoTen) continue;
    if (trangThai === "Đã lưu mới" || trangThai === "Đã cập nhật") continue;
    
    try {
      var contact = null;
      var cccdKey = "CCCD: " + soCCCD;
      
      // 1. Tìm trùng bằng CCCD trong danh bạ
      if (soCCCD) {
        var foundContacts = ContactsApp.searchContacts(cccdKey);
        for (var j = 0; j < foundContacts.length; j++) {
          if (foundContacts[j].getNotes().indexOf(cccdKey) !== -1) {
            contact = foundContacts[j];
            break;
          }
        }
      }
      
      // 2. Nội dung ghi chú
      var chuoiGhiChuMoi = "";
      if (maNV) chuoiGhiChuMoi += "Mã NV: " + maNV + "\n";
      chuoiGhiChuMoi += cccdKey;
      if (congTy) chuoiGhiChuMoi += "\nNơi công tác: " + congTy;
      if (boPhan) chuoiGhiChuMoi += "\nBộ phận: " + boPhan;
      if (chucVu) chuoiGhiChuMoi += "\nChức vụ: " + chucVu;
      
      if (contact) {
        // === CẬP NHẬT BIẾN ĐỘNG ===
        contact.setFullName(hoTen);
        contact.setNotes(chuoiGhiChuMoi);
        
        if (soDienThoai) {
          var phones = contact.getPhones();
          var isPhoneExist = phones.some(function(p) {
            return p.getPhoneNumber().replace(/[\s\-\(\)\.]/g, "") === soDienThoai.replace(/[\s\-\(\)\.]/g, "");
          });
          if (!isPhoneExist) {
            contact.addPhone(ContactsApp.Field.MAIN_PHONE, soDienThoai);
          }
        }
        
        if (email) {
          var emails = contact.getEmails();
          emails.forEach(function(e) { e.deleteEmailField(); });
          contact.addEmail(ContactsApp.Field.HOME_EMAIL, email);
        }
        
        var addresses = contact.getAddresses();
        addresses.forEach(function(a) { a.deleteAddressField(); });
        if (thuongTru) contact.addAddress(ContactsApp.Field.HOME_ADDRESS, thuongTru);
        if (tamTru) contact.addAddress(ContactsApp.Field.WORK_ADDRESS, tamTru);
        
        if (ngaySinh && ngaySinh instanceof Date) {
          updateBirthday(contact, ngaySinh);
        }
        
        sheet.getRange(i + 2, 13).setValue("Đã cập nhật");
        countUpdated++;
        
      } else {
        // === TẠO MỚI ===
        contact = ContactsApp.createContact(hoTen, '', email);
        contact.setNotes(chuoiGhiChuMoi);
        
        if (soDienThoai) contact.addPhone(ContactsApp.Field.MAIN_PHONE, soDienThoai);
        if (thuongTru) contact.addAddress(ContactsApp.Field.HOME_ADDRESS, thuongTru);
        if (tamTru) contact.addAddress(ContactsApp.Field.WORK_ADDRESS, tamTru);
        
        if (ngaySinh && ngaySinh instanceof Date) {
          updateBirthday(contact, ngaySinh);
        }
        
        sheet.getRange(i + 2, 13).setValue("Đã lưu mới");
        countCreated++;
      }
      
      // Xử lý nhãn
      if (labelName && contact) {
        if (!labelCache[labelName]) {
          var group = ContactsApp.getContactGroup(labelName);
          if (!group) group = ContactsApp.createContactGroup(labelName);
          labelCache[labelName] = group;
        }
        labelCache[labelName].addContact(contact);
      }
      
    } catch (e) {
      Logger.log("Lỗi ở dòng " + (i + 2) + ": " + e.message);
      sheet.getRange(i + 2, 13).setValue("Lỗi: " + e.message);
    }
  }
  
  SpreadsheetApp.getUi().alert('Kết quả', 'Đã xử lý xong!\n- Tạo mới: ' + countCreated + ' nhân viên.\n- Cập nhật: ' + countUpdated + ' nhân viên.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// HÀM MỚI: Tự động tách chuỗi Ngày/Tháng/Năm chuẩn Việt Nam để nạp vào Google Contacts
function updateBirthday(contact, ngaySinhRaw) {
  try { 
    // Xóa ngày sinh cũ nếu có để tránh bị rác dữ liệu
    contact.getBirthday().deleteBirthdayField(); 
  } catch(e) {}
  
  if (!ngaySinhRaw) return;
  
  var ngay, thang, nam;
  
  // Trường hợp 1: Nếu Google Sheets trả về dạng chuỗi văn bản (Text) như "28/04/1985"
  if (typeof ngaySinhRaw === "string" || ngaySinhRaw instanceof String) {
    var parts = ngaySinhRaw.split("/");
    if (parts.length === 3) {
      ngay = parseInt(parts[0], 10);
      thang = parseInt(parts[1], 10);
      nam = parseInt(parts[2], 10);
    } else {
      return; // Không đúng định dạng ngày/tháng/năm thì bỏ qua
    }
  } 
  // Trường hợp 2: Nếu ô đó đã được Google Sheets hiểu sẵn là định dạng Ngày Tháng (Date Object)
  else if (ngaySinhRaw instanceof Date) {
    ngay = ngaySinhRaw.getDate();
    thang = ngaySinhRaw.getMonth() + 1; // Tháng trong JS bắt đầu từ 0
    nam = ngaySinhRaw.getFullYear();
  } else {
    return;
  }
  
  // Kiểm tra tính hợp lệ của số liệu sau khi tách
  if (isNaN(ngay) || isNaN(thang) || isNaN(nam) || thang < 1 || thang > 12) return;
  
  // Chuyển đổi tháng sang dạng chữ tiếng Anh viết tắt chuẩn của Google API
  var months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
  var monthEnum = ContactsApp.Month[months[thang - 1]];
  
  // Nạp chính xác vào danh bạ Google
  contact.setBirthday(monthEnum, ngay, nam);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Công cụ Danh bạ nâng cao')
    .addItem('Đồng bộ dữ liệu tổng hợp', 'dongBoDanhBaTongHop')
    .addToUi();
}
