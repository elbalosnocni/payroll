# Payroll Portal – FINAL Optimized 2026

## Bộ file

- `gas/Code.gs` — Google Apps Script API/backend
- `web/index.html` — GitHub Pages UI
- `web/app.js` — frontend logic
- `web/styles.css` — responsive UI
- `vba/SyncPayroll.bas` — đồng bộ Excel XLSB → Google Sheets
- `README.md` — hướng dẫn triển khai

## Các lỗi đã kiểm tra và sửa

### 1. VBA vô tình đọc cả dòng tiêu đề
Bản cũ bắt đầu vòng lặp từ row 1, nên header `Mã NV / Họ tên` có thể bị xem như một nhân viên.

**Bản FINAL:** Salary và DSCNV đều xử lý dữ liệu từ row 2.

### 2. Đồng bộ batch có nguy cơ sai dữ liệu khi retry / chạy song song
Bản cũ chỉ dùng `LockService` trong từng request. Lock không tồn tại xuyên suốt toàn bộ chuỗi batch.

**Bản FINAL:** GAS có trạng thái sync theo `Factory + PayMonth`:
- Batch 1 chỉ clear dữ liệu đúng kỳ đúng Factory một lần.
- Batch sau phải đi đúng thứ tự.
- Retry cùng `BatchID` không clear lại dữ liệu.
- Không cho một `BatchID` khác ghi đè một sync đang chạy.
- Có TTL trạng thái để tránh bị khóa vĩnh viễn nếu máy đồng bộ chết giữa chừng.

### 3. Clear dữ liệu cũ tạo rất nhiều dòng trống
Bản cũ dùng `clearContent()` nên sau nhiều lần đồng bộ, Sheet Salary có thể phình ra rất nhiều dòng rỗng.

**Bản FINAL:** xóa các dòng thuộc đúng `Factory + PayMonth` từ dưới lên, sau đó batch mới ghi lại gọn hơn.

### 4. CCCD / EmployeeID phải giữ số 0 đầu
GAS đặt các cột định danh thành Text:
- Users: A:B
- Salary: D:F

VBA vẫn đọc CCCD từ DSCNV H và loại bỏ dấu `'` đầu chuỗi trước khi gửi.

**Lưu ý quan trọng:** nếu file nguồn DSCNV đã lưu CCCD dạng số và Excel đã làm mất số 0 từ trước thì VBA không thể khôi phục số 0 đó. Tốt nhất cột H của DSCNV phải là Text.

### 5. Không cho chạy sync nếu vẫn dùng API key placeholder
VBA sẽ dừng trước khi gửi nếu còn:
`PUT_YOUR_RANDOM_SYNC_API_KEY_HERE`

### 6. JSON từ VBA an toàn hơn
Đã bổ sung xử lý thêm các ký tự điều khiển JSON như Backspace và Form Feed.

### 7. Setup Google Sheet
Đã thêm:
`setupSystem()`

Chạy hàm này một lần sau khi cài Code.gs để tự:
- tạo Users / Salary / SyncLog nếu thiếu
- tạo/sửa header
- định dạng các cột ID dạng Text
- freeze dòng header
- auto resize cột

### 8. Payslip vẫn kiểm tra quyền ở server
Frontend không thể tự sửa EmployeeID để xem lương người khác. GAS lấy EmployeeID từ User đang đăng nhập rồi mới truy vấn Salary.

### 9. Bắt buộc đổi mật khẩu lần đầu / sau Admin reset
Logic này vẫn được giữ và kiểm tra ở server.

---

# Cài đặt GAS

## 1. Dán Code.gs

Mở Google Apps Script và thay toàn bộ `Code.gs` bằng file:

`gas/Code.gs`

## 2. Script Properties

Vào:

**Project Settings → Script Properties**

Tạo:

- `SPREADSHEET_ID` = ID Google Sheet chứa payroll
- `SYNC_API_KEY` = một chuỗi ngẫu nhiên dài, ví dụ tối thiểu 32 ký tự

Không dùng API key mẫu trong README/VBA.

## 3. Chạy setupSystem()

Trong Apps Script chọn:

`setupSystem`

và Run một lần.

## 4. Tạo mật khẩu Admin

Chạy:

`setAdminPassword("MAT_KHAU_ADMIN_MOI")`

Ví dụ mật khẩu thực tế phải do bạn tự đặt.

## 5. Deploy Web App

- Execute as: **Me**
- Who has access: **Anyone**

Sau khi deploy, copy URL `/exec`.

Nếu URL deployment thay đổi, sửa cả:

- `web/app.js`
- `vba/SyncPayroll.bas`

---

# Google Sheets

## Users

Các cột:

`Username | EmployeeID | FullName | PasswordHash | PasswordSalt | MustChangePassword | Active | UpdatedAt`

Username hiện được thiết kế là **CCCD**.

## Salary

Không tự ý đổi thứ tự 44 cột mà Code.gs đang sử dụng.

Khóa dữ liệu:

`Factory + PayMonth + EmployeeID`

## SyncLog

Ghi từng batch:

`UpdatedAt | Factory | PayMonth | BatchID | BatchNo | TotalBatches | Rows | Inserted | Updated | Status | Message`

---

# VBA

Mở VBA Editor → Import:

`vba/SyncPayroll.bas`

Sửa:

```vb
Private Const SYNC_API_KEY As String = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE"
```

thành API key thật giống `SYNC_API_KEY` trong Script Properties.

## Macro

- `SyncBothFactories` — đồng bộ Snack + Flexible
- `SyncSnackOnly` — chỉ Snack
- `SyncFlexibleOnly` — chỉ Flexible

Kỳ lương mặc định là **tháng trước**.

Ví dụ ngày 10/09/2026 → `08-2026`.

---

# Nguồn dữ liệu

### Snack

`\\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW\YYYY\SALARY MM-YYYY.xlsb`

### Flexible

`\\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line\YYYY\PRINTING LINE MM-YYYY.xlsb`

Password XLSB hiện giữ theo bản gốc:

`1234`

Nếu password file nguồn thay đổi, sửa `XLS_PASSWORD` trong VBA.

---

# DSCNV mapping

- D = Mã NV
- B = Họ tên
- H = CCCD
- AF = Phòng ban
- AG = Bộ phận
- AH = Chức vụ

Join:

1. Mã NV trước
2. Họ tên làm fallback

---

# Kiểm tra trước khi chạy thật

1. Chạy `setupSystem()`.
2. Kiểm tra Users / Salary / SyncLog.
3. Đặt `SYNC_API_KEY`.
4. Đặt Admin password.
5. Deploy lại Web App.
6. Cập nhật GAS URL trong `app.js` và VBA.
7. Chạy `SyncSnackOnly` với một kỳ test.
8. Kiểm tra SyncLog.
9. Kiểm tra Salary không có header bị import thành nhân viên.
10. Kiểm tra CCCD bắt đầu bằng `0`.
11. Đăng nhập bằng CCCD.
12. Kiểm tra chỉ xem được phiếu lương của chính mình.
13. Thử retry / chạy lại cùng kỳ để kiểm tra không bị nhân đôi.

## Security

API key nằm trong VBA nên người có quyền xem VBA vẫn có thể lấy key. Vì vậy đây là **sync gate**, không phải cơ chế bí mật tuyệt đối.

Nếu hệ thống được đưa vào production quy mô lớn, nên chuyển quá trình sync sang một service/server được kiểm soát thay vì phân phối API key trong nhiều file Excel.

## Phiên bản

`Payroll Portal – FINAL Optimized 2026`

Mục tiêu: ổn định hơn khi sync batch, tránh header giả, giảm dòng trống, bảo toàn ID dạng text và kiểm soát thứ tự batch.
