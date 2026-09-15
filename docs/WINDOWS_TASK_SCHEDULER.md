# Lên lịch chạy tự động (Windows Task Scheduler)

Mục tiêu: vào ngày 10 hằng tháng, tự động mở Excel, chạy macro
`SyncPayroll.RunSync`, đẩy dữ liệu lương tháng trước lên Google Sheet.

## 1. Chuẩn bị file Excel chứa macro

1. Tạo 1 file `.xlsm` (Excel Macro-Enabled Workbook) riêng, ví dụ
   `PayrollSyncRunner.xlsm` — **không** nhúng macro trực tiếp vào các file
   lương `.xlsb` gốc.
2. Mở VBA Editor (Alt+F11), import 3 module trong thư mục `vba/`:
   `Config.bas`, `Json.bas`, `SyncPayroll.bas` (File > Import File...).
3. Sửa `Config.bas`:
   - `GAS_URL`: URL Web App đã deploy (xem `GAS_DEPLOY.md`).
   - `SYNC_API_KEY`: phải khớp với Script Properties bên Apps Script.
   - `XLSB_PASSWORD`, các đường dẫn `SNACK_ROOT` / `FLEXIBLE_ROOT` nếu có
     thay đổi.
4. Lưu file `.xlsm`, để ở vị trí cố định, ví dụ:
   `C:\PayrollSync\PayrollSyncRunner.xlsm`

## 2. Tạo file VBS gọi Excel chạy macro (không cần mở giao diện)

Tạo file `RunSync.vbs` với nội dung:

```vbscript
Dim excelApp, workbook
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
excelApp.DisplayAlerts = False

Set workbook = excelApp.Workbooks.Open("C:\PayrollSync\PayrollSyncRunner.xlsm")
excelApp.Run "SyncPayroll.RunSync"

workbook.Close False
excelApp.Quit

Set workbook = Nothing
Set excelApp = Nothing
```

## 3. Tạo Task trong Task Scheduler

1. Mở **Task Scheduler** > **Create Task...** (không dùng "Create Basic
   Task" để có đủ tuỳ chọn nâng cao).
2. Tab **General**:
   - Đặt tên: `Payroll Monthly Sync`.
   - Chọn **Run whether user is logged on or not**.
   - Chọn **Run with highest privileges** nếu cần quyền truy cập share
     `\\192.168.0.253\...`.
3. Tab **Triggers** > **New**:
   - **Begin the task**: On a schedule.
   - **Monthly**, chọn ngày **10**, tất cả các tháng.
   - Giờ chạy: ví dụ 08:00 sáng (đảm bảo file lương tháng trước đã được
     hoàn thiện và lưu trên share trước giờ này).
4. Tab **Actions** > **New**:
   - **Action**: Start a program.
   - **Program/script**: `wscript.exe`
   - **Add arguments**: `"C:\PayrollSync\RunSync.vbs"`
5. Tab **Conditions**: bỏ chọn "Start the task only if the computer is on
   AC power" nếu chạy trên máy bàn luôn cắm điện.
6. Tab **Settings**: bật "Run task as soon as possible after a scheduled
   start is missed" để phòng trường hợp máy tắt đúng lúc lịch chạy.
7. Lưu Task, nhập mật khẩu tài khoản Windows có quyền truy cập share mạng.

## 4. Kiểm tra log

Macro ghi log vào `%TEMP%\SyncPayroll.log` (xem hàm `LogMessage` trong
`SyncPayroll.bas`) và vào sheet `SyncStatus` trên Google Sheet (xem trong
trang Admin của web app, tab "Trạng thái đồng bộ"). Nếu file không được
tìm thấy hoặc mở file lỗi (sai mật khẩu, đường dẫn không tồn tại...), log sẽ
ghi rõ nguyên nhân.

## 5. Chạy thử thủ công trước khi đưa vào lịch

Mở `PayrollSyncRunner.xlsm`, nhấn Alt+F8, chọn `SyncPayroll.RunSync`, bấm
Run — theo dõi cửa sổ Immediate (Ctrl+G trong VBA Editor) và file log để
xác nhận mọi thứ chạy đúng trước khi giao cho Task Scheduler tự động hoá.
