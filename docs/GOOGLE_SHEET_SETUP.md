# Thiết lập Google Sheet

## 1. Tạo Spreadsheet

1. Tạo một Google Sheet mới, đặt tên ví dụ `Payroll Database`.
2. Lấy **Spreadsheet ID** từ URL:
   `https://docs.google.com/spreadsheets/d/`**`SPREADSHEET_ID`**`/edit`

## 2. Khởi tạo cấu trúc sheet

1. Mở **Extensions > Apps Script** từ chính Spreadsheet này (để Apps Script
   gắn liền với Sheet, không cần cấu hình thêm quyền).
2. Dán toàn bộ code trong thư mục `gas/` vào các file `.gs` tương ứng
   (tạo file mới trong Apps Script Editor với đúng tên: `Code.gs`,
   `Config.gs`, `Auth.gs`, `Payroll.gs`, `Admin.gs`, `Sync.gs`, `Utils.gs`).
3. Vào **Project Settings > Script Properties**, thêm các key:
   - `SPREADSHEET_ID` = ID lấy ở bước 1 (nếu chạy Apps Script gắn liền với
     Sheet, có thể để trống và sửa `Config.gs` dùng `SpreadsheetApp.getActiveSpreadsheet()`
     thay vì `openById`).
   - `SYNC_API_KEY` = một chuỗi bí mật tự đặt, dùng để VBA xác thực khi đẩy
     dữ liệu lên (phải khớp với `SYNC_API_KEY` trong `vba/Config.bas`).
   - `PASSWORD_PEPPER` = một chuỗi bí mật khác, dùng để băm mật khẩu nhân
     viên an toàn hơn.
4. Chạy hàm `setupSheets` (chọn hàm trong dropdown, bấm Run) để tự động tạo
   4 sheet: `Employees`, `Payroll`, `AuditLog`, `SyncStatus` với đầy đủ cột.
5. Cấp quyền chạy Script khi được yêu cầu (chọn tài khoản Google, bấm
   "Advanced" > "Go to ... (unsafe)" nếu màn hình cảnh báo hiện ra — bình
   thường vì đây là script tự viết của bạn).

## 3. Cột CCCD giữ số 0 ở đầu

Cột `CCCD` trong sheet `Employees` cần được định dạng là **Plain text**
(chọn cột > Format > Number > Plain text) **trước khi** có dữ liệu, để tránh
Google Sheets tự động cắt số 0 ở đầu. Code trong `Sync.gs` ghi giá trị CCCD
kèm dấu `'` ở đầu (ví dụ `'0079012345`) để ép Google Sheets hiểu là văn bản,
nhưng đặt sẵn định dạng cột là Plain text vẫn là cách an toàn nhất.

## 4. Tạo tài khoản Admin đầu tiên

1. Mở sheet `Employees`, thêm 1 dòng thủ công:
   - `CCCD`: số CCCD của admin (dạng text)
   - `MaNV`: mã nhân viên bất kỳ, ví dụ `ADMIN01`
   - `HoTen`, `Xuong`, `PhongBan`, `BoPhan`, `ChucVu`: tuỳ chọn
   - `Role`: `admin`
   - `MustChangePassword`: `TRUE`
   - Để trống `PasswordHash`, `PasswordSalt`
2. Trong Apps Script Editor, mở tab Apps Script, chạy tạm đoạn code sau
   (Run > Run function, hoặc gõ vào 1 hàm test rồi Run) để đặt mật khẩu ban
   đầu cho admin bằng chính Mã nhân viên (giống quy trình reset password):

   ```javascript
   function initAdminPassword() {
     var sheet = getOrCreateSheet(CONFIG.SHEET_EMPLOYEES);
     var rows = sheetToObjects(sheet);
     var admin = rows.filter(function(r){ return r.MaNV === 'ADMIN01'; })[0];
     var salt = generateSalt();
     var hash = hashPassword(admin.MaNV, salt);
     var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
     sheet.getRange(admin.__row, headers.indexOf('PasswordHash')+1).setValue(hash);
     sheet.getRange(admin.__row, headers.indexOf('PasswordSalt')+1).setValue(salt);
   }
   ```

3. Đăng nhập lần đầu bằng CCCD + mật khẩu = `ADMIN01`, hệ thống sẽ bắt buộc
   đổi mật khẩu ngay.

Sau đó, mọi thao tác reset mật khẩu tiếp theo (kể cả cho admin khác) có thể
thực hiện trực tiếp từ trang Admin trong web app.
