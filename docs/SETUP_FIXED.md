# Cài đặt bản FIXED

## 1. Google Sheet

Tạo một Google Spreadsheet riêng cho hệ thống. Không cần tự tạo sheet; GAS sẽ tạo:

- Employees
- Payroll
- AuditLog
- SyncStatus

## 2. Apps Script

Mở Apps Script project làm API và thay toàn bộ file trong thư mục `gas/`.

Trong **Project Settings > Script properties** tạo đúng 3 property:

| Property | Giá trị |
|---|---|
| `SPREADSHEET_ID` | ID của Google Sheet database |
| `SYNC_API_KEY` | Một chuỗi bí mật tự đặt, ví dụ 32+ ký tự |
| `PASSWORD_PEPPER` | Một chuỗi bí mật khác, ví dụ 32+ ký tự |

Không dùng `CHANGE_ME...`.

Chạy hàm `setupSheets()` một lần trong Apps Script Editor.

### Kiểm tra GAS

Mở URL Web App bằng trình duyệt. Phải nhận JSON có `ok: true` và `Payroll API đang hoạt động.`

Có thể gọi action `health` từ frontend/VBA sau khi cấu hình để kiểm tra Spreadsheet ID và danh sách sheet.

## 3. Deploy Web App

Deploy > New deployment > Web app:

- Execute as: **Me**
- Who has access: **Anyone**

Sau khi sửa code GAS, phải **Deploy > Manage deployments > Edit > New version > Deploy**. GitHub Pages phải dùng đúng URL `/exec` của deployment đang chạy.

## 4. VBA

Import 3 module:

- `vba/Config.bas`
- `vba/Json.bas`
- `vba/SyncPayroll.bas`

Trong `Config.bas`, sửa:

```vb
Public Const SYNC_API_KEY As String = "..."
```

cho **trùng tuyệt đối** với Script Property `SYNC_API_KEY` của GAS.

Các đường dẫn UNC và password XLSB đã đặt theo yêu cầu hiện tại.

## 5. Chạy test

Trong Excel VBA chạy:

```text
RunSyncForSpecificMonth
```

Mặc định hàm test đang lấy `08-2026`.

Khi chạy thực tế, dùng:

```text
RunSync
```

Hàm tự lấy tháng trước tháng hiện tại. Ví dụ 14/09/2026 -> `08-2026`.

## 6. Xem log VBA

Mở:

```text
%TEMP%\SyncPayroll.log
```

Bản FIXED không còn `On Error Resume Next` bao trùm toàn bộ macro. Log phải cho biết:

- đường dẫn file
- mở XLSB thành công
- tìm thấy Salary/DSCNV
- số dòng
- số employee
- số payroll
- JSON length
- HTTP Status
- response từ GAS

Nếu Google Sheet vẫn trắng, **đây là file log quan trọng nhất**.

## 7. GitHub Pages

Trong `web/auth.js`, `GAS_URL` đã đặt URL Web App hiện tại được cung cấp trong yêu cầu.

Nếu sau này URL deployment thay đổi, sửa đúng dòng `GAS_URL`.

## 8. Tạo Admin đầu tiên

Sau khi `Employees` có dữ liệu, sửa trực tiếp một nhân viên thành:

```text
Role = admin
```

Tài khoản nhân viên mới có mật khẩu ban đầu = `MaNV` và `MustChangePassword = TRUE`.

Admin có thể reset mật khẩu về `MaNV`.

## 9. Không bỏ qua đổi mật khẩu

Bản FIXED đã bỏ nút "Bỏ qua". Khi `MustChangePassword = TRUE`, nhân viên bắt buộc hoàn thành đổi mật khẩu trước khi sử dụng hệ thống bình thường.
