# Payroll GitHub Pages + Google Apps Script

Hệ thống tra cứu phiếu lương cho nhân viên, kiến trúc:

Excel XLSB → VBA Runner → GAS `/sync` → Google Sheets → GitHub Pages → Login CCCD/password → session → chỉ dữ liệu của chính nhân viên.

## Bảo mật
- CCCD là định danh đăng nhập, không được gửi `MaNV` từ frontend khi đọc lương.
- Backend lấy `MaNV` từ session rồi lọc Payroll theo `MaNV`.
- Mật khẩu không lưu dạng rõ: `PasswordHash` + `PasswordSalt`.
- Hash SHA-256 lặp nhiều vòng + `PASSWORD_PEPPER` trong Script Properties.
- Nhân viên mới/reset: password = MaNV, `MustChangePassword=true`.
- Sau đổi mật khẩu, backend cấp token mới và hủy token cũ.
- Admin được kiểm tra quyền ở backend, không chỉ ẩn nút frontend.
- Login có rate-limit cơ bản.
- AuditLog ghi login, đổi mật khẩu, xem lương, reset, sync.
- CCCD luôn được lưu Plain text.

> Không commit `SYNC_API_KEY` thật hoặc `PASSWORD_PEPPER` lên GitHub.

## 1. Google Apps Script
Tạo Apps Script gắn với Spreadsheet hoặc standalone. Import:
`gas/Config.gs`, `Code.gs`, `Utils.gs`, `Auth.gs`, `Payroll.gs`, `Sync.gs`, `Admin.gs`.

Script Properties:
- `SPREADSHEET_ID`
- `SYNC_API_KEY`
- `PASSWORD_PEPPER`

Chạy `setupSheets()` một lần.

Tạo admin trong Employees, ví dụ Role=`ADMIN`, sau đó chạy:
`setInitialAdmin("ADMIN01")`

## 2. Deploy GAS
Deploy Web app:
- Execute as: Me
- Who has access: Anyone

URL hiện tại của project đã được điền trong `web/auth.js`. Nếu deployment đổi URL, cập nhật `auth.js` và `vba/Config.bas`.

## 3. VBA
Import:
- `Config.bas`
- `Json.bas`
- `SyncPayroll.bas`

Cấu trúc file:
- Snack: `<root>\2026\SALARY 08-2026.xlsb`
- Flexible: `<root>\2026\PRINTING LINE 08-2026.xlsb`

Mật khẩu XLSB: `1234`.

Macro `RunSync` lấy tháng trước. Macro `RunScheduledSync` chỉ chạy vào ngày 10.

## 4. Task Scheduler
Đặt Task chạy ngày 10 hàng tháng, ví dụ 08:00, gọi `RunScheduledSync` qua `RunSync.vbs`.
Runner nên nằm trên máy Windows có quyền đọc UNC share và có Excel.

## 5. GitHub Pages
Đưa toàn bộ thư mục `web/` lên root của repository:
- `index.html`
- `auth.js`
- `app.js`
- `admin.js`
- `styles.css`

Bật Settings → Pages → Deploy from branch.

## 6. Dữ liệu
Salary và DSCNV đọc từ dòng 7.
Join DSCNV theo Họ tên sau khi chuẩn hóa bỏ dấu/khoảng trắng; MaNV là khóa chính của Payroll.

UpdatedAt được GAS ghi theo:
`dd/MM/yyyy HH:mm:ss`

Lưu ý: Apps Script không thể trực tiếp đọc UNC `\\192.168...`; vì vậy VBA trên Windows là thành phần đọc XLSB và đẩy JSON lên GAS.
