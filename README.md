# Payroll Portal – Optimized 2026 (FIXED)

Bản này đã rà soát toàn bộ `GAS + GitHub Pages + VBA` và sửa các lỗi quan trọng trong ZIP gốc.

## Các lỗi quan trọng đã sửa

### 1. Lỗi lớn: Users không được tạo/cập nhật khi đồng bộ Payroll
ZIP gốc có hàm `syncUsers_()` nhưng không gọi hàm này trong `syncPayroll_()`.

Hậu quả:
- Salary có dữ liệu nhưng Sheet `Users` có thể không có nhân viên.
- Nhân viên mới không đăng nhập được.
- Mã nhân viên/CCCD mới không tự đồng bộ sang Users.

**Bản FIXED:** mỗi batch Payroll tự đồng bộ Users từ `CitizenID + EmployeeID + FullName`.
- Nhân viên mới: tạo User.
- Nhân viên cũ: cập nhật Mã NV/Họ tên/Active.
- **Không ghi đè PasswordHash/PasswordSalt**, nên mật khẩu nhân viên đã đổi vẫn được giữ.
- Đồng bộ lặp lại vẫn an toàn.

### 2. CCCD giữ số 0 đầu
GAS định dạng `Users A:B` và `Salary D:F` là Text.

VBA:
- loại dấu `'` đầu CCCD;
- nếu nguồn vô tình biến CCCD thành số 11 chữ số, bản FIXED tự thêm `0` phía trước;
- tốt nhất vẫn phải giữ cột H của `DSCNV` là Text.

> Nếu Excel đã mất nhiều chữ số do số quá dài/định dạng sai trước khi VBA đọc, không thể khôi phục chính xác bằng phần mềm.

### 3. Đăng nhập không lưu thông tin cũ trên giao diện
GitHub Pages không lưu token/localStorage.

Bản FIXED:
- xóa CCCD/mật khẩu khi logout;
- xóa form khi tải trang;
- xử lý `pageshow` để tránh dữ liệu cũ từ browser back-forward cache;
- tắt autocomplete ở các ô nhạy cảm.

### 4. Loading rõ ràng
Các nút:
- Đăng nhập
- Đổi mật khẩu
- Xem phiếu
- Đăng nhập Admin
- Reset mật khẩu

đều có spinner + trạng thái disabled trong lúc gọi API.

### 5. Hiển thị kỳ lương
Giá trị API vẫn là `MM-YYYY`.

Giao diện hiển thị thân thiện:
`Tháng 08/2026`

Ngày 10/09/2026 chạy VBA sẽ tự tính:
`08-2026`.

### 6. Join Salary ↔ DSCNV
Ưu tiên:
1. Mã nhân viên `D`
2. Nếu không tìm thấy, dùng Họ tên `B`

DSCNV:
- B = Họ tên
- D = Mã NV
- H = CCCD
- AF = Phòng ban
- AG = Bộ phận
- AH = Chức vụ

Salary:
- D = Mã NV
- CF = Họ tên

### 7. Batch sync
GAS vẫn giữ cơ chế:
- Batch 1 chỉ clear đúng `Factory + PayMonth`.
- Batch sau phải đi đúng thứ tự.
- Một `BatchID` khác không được ghi đè sync đang chạy.
- Có TTL chống khóa vĩnh viễn.
- Retry cùng batch là an toàn theo trạng thái sync.

---

# 1. Cấu trúc

```text
Payroll_Portal_Optimized_2026_FIXED/
├── README.md
├── gas/
│   └── Code.gs
├── web/
│   ├── index.html
│   ├── app.js
│   └── styles.css
└── vba/
    └── SyncPayroll.bas
```

---

# 2. Google Apps Script

Dán toàn bộ:

`gas/Code.gs`

vào Apps Script.

## Script Properties

Tạo:

```text
SPREADSHEET_ID = ID Google Sheet
SYNC_API_KEY   = chuỗi ngẫu nhiên tối thiểu 32 ký tự
```

**Không dùng API key đang nằm trong ZIP cũ.**

API key trong VBA có thể bị người có quyền xem file Excel/VBA đọc được, vì vậy đây chỉ là lớp bảo vệ endpoint đồng bộ, không phải bí mật tuyệt đối.

## Setup

Chạy một lần:

```text
setupSystem()
```

Hệ thống tự tạo:

```text
Users
Salary
SyncLog
```

## Tạo Admin

Chạy:

```text
setAdminPassword("MAT_KHAU_ADMIN_CUA_BAN")
```

Không ghi mật khẩu thật vào README hoặc mã nguồn.

---

# 3. Deploy GAS

Deploy Web App:

- Execute as: **Me**
- Who has access: **Anyone**

Sau đó lấy URL `/exec`.

Cập nhật URL tại:

```text
web/app.js
vba/SyncPayroll.bas
```

**Lưu ý:** URL GAS trong ZIP này đang giữ URL bạn đã cung cấp. Nếu bạn redeploy GAS và URL thay đổi, phải sửa cả 2 file trên.

---

# 4. Google Sheets

## Users

```text
Username
EmployeeID
FullName
PasswordHash
PasswordSalt
MustChangePassword
Active
UpdatedAt
```

Username = CCCD.

Mật khẩu:
- lần đầu = Mã NV;
- sau khi đổi = SHA-256 với salt riêng;
- Admin reset = Mã NV + bắt buộc đổi lại.

## Salary

44 cột theo `SALARY_HEADERS` trong `Code.gs`.

Khóa logic:

```text
Factory + PayMonth + EmployeeID
```

## SyncLog

Ghi từng batch:

```text
UpdatedAt
Factory
PayMonth
BatchID
BatchNo
TotalBatches
Rows
Inserted
Updated
Status
Message
```

`UpdatedAt` dùng:

```text
dd/MM/yyyy HH:mm:ss
```

Timezone:

```text
Asia/Ho_Chi_Minh
```

---

# 5. Nguồn Excel

## Snack

```text
\192.168.0.253n hr\SALARY - 2014 - 2015\VNLWW\YYYY\SALARY MM-YYYY.xlsb
```

## Flexible

```text
\192.168.0.253n hr\SALARY - 2014 - 2015\Printing line\YYYY\PRINTING LINE MM-YYYY.xlsb
```

Password XLSB:

```text
1234
```

VBA tự lấy **tháng trước**.

Ví dụ:

```text
10/09/2026
    ↓
08-2026
    ↓
2026\SALARY 08-2026.xlsb
2026\PRINTING LINE 08-2026.xlsb
```

---

# 6. Macro VBA

Import:

```text
vba/SyncPayroll.bas
```

Sửa:

```vb
Private Const SYNC_API_KEY As String = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE"
```

thành đúng giá trị `SYNC_API_KEY` trong Script Properties.

## Macro chính

```text
SyncBothFactories
```

Đồng bộ:
- Snack
- Flexible

Macro riêng:

```text
SyncSnackOnly
SyncFlexibleOnly
```

Macro kiểm tra đường dẫn:

```text
CheckPreviousPayrollFiles
```

Macro này không upload dữ liệu; chỉ kiểm tra file kỳ lương tháng trước có tồn tại hay không.

---

# 7. Dữ liệu Salary

## Thông tin nhân viên

```text
CF = Họ tên
D  = Mã NV
```

## Thu nhập

```text
AR = Tổng các khoản thu nhập
AN = Lương tháng
F  = Lương cơ bản
X  = Tiền khác
Y  = Tiền kỷ luật
Z  = Gắn bó 2 năm
AA = Gắn bó 5 năm
AB = Gắn bó 10 năm
AC = Tiền nhà ở
AD = Tiền đi lại
AE = Thưởng chuyên cần
AQ = Hoa hồng / thưởng vượt định mức
AF = Trợ cấp thôi việc + phép năm còn lại
```

## Tăng ca

```text
AO = Lương làm ngoài giờ/ngày nghỉ/ca đêm
L  = Giờ ngoài giờ
M  = Giờ ngày nghỉ
N  = Giờ ngoài giờ ngày nghỉ
P  = Giờ tăng ca đêm ngày nghỉ
R  = Giờ ngày lễ
S  = Giờ ngoài giờ ngày lễ
T  = Giờ tăng ca đêm ngày lễ
U  = Ngày làm ca đêm
Q  = Giờ tăng ca đêm
```

## Khấu trừ

```text
AG = BHXH
AH = BHYT
AI = BHTN
AT = Thuế TNCN
AK = Tạm ứng
AJ = Khấu trừ khác
```

## Thực lĩnh

```text
AW = Lương thực lĩnh
```

---

# 8. Định dạng

Tiền:
- hiển thị VND theo `vi-VN`;
- có dấu phân cách hàng nghìn.

Ngày/giờ:
```text
dd/MM/yyyy HH:mm:ss
```

Số ngày / số giờ:
- tối đa 2 số thập phân.

CCCD:
- xử lý dạng Text;
- không tự bỏ số 0 đầu.

---

# 9. Luồng đăng nhập

```text
CCCD + mật khẩu
       ↓
GAS kiểm tra Users
       ↓
Mật khẩu lần đầu = Mã NV
       ↓
mustChangePassword = TRUE
       ↓
bắt buộc đổi mật khẩu
       ↓
SHA-256 + Salt
       ↓
PasswordHash / PasswordSalt
       ↓
xem phiếu lương
```

Sau Admin Reset:

```text
Admin reset
    ↓
Password = Mã NV
    ↓
MustChangePassword = TRUE
    ↓
nhân viên đăng nhập
    ↓
bắt buộc đổi mật khẩu
```

---

# 10. Bảo mật

Frontend **không quyết định quyền xem lương**.

GAS lấy:

```text
token → Users → EmployeeID
```

sau đó mới tìm:

```text
PayMonth + EmployeeID
```

Do đó nhân viên không thể sửa JavaScript để truyền EmployeeID của người khác rồi xem phiếu.

Không lưu:
- CCCD
- mật khẩu
- token

vào localStorage.

---

# 11. Kiểm tra trước khi chạy thật

Theo thứ tự:

1. Dán `Code.gs`.
2. Tạo Script Properties.
3. Chạy `setupSystem()`.
4. Chạy `setAdminPassword(...)`.
5. Deploy Web App.
6. Cập nhật GAS URL ở `app.js` và `SyncPayroll.bas`.
7. Đặt `SYNC_API_KEY` thật.
8. Import `SyncPayroll.bas`.
9. Chạy `CheckPreviousPayrollFiles`.
10. Chạy `SyncSnackOnly`.
11. Kiểm tra `SyncLog`.
12. Kiểm tra `Salary`.
13. Kiểm tra `Users`.
14. Kiểm tra CCCD có số 0 đầu.
15. Đăng nhập bằng CCCD + Mã NV.
16. Đổi mật khẩu.
17. Đăng xuất.
18. Đăng nhập lại bằng mật khẩu mới.
19. Kiểm tra chỉ xem được phiếu của chính mình.
20. Admin reset thử một tài khoản test.

---

# 12. Lưu ý về ngày 10

VBA tự tính **tháng trước**, không phụ thuộc việc máy chạy đúng ngày 10.

Vì vậy:

```text
09/09/2026 → kỳ 08-2026
10/09/2026 → kỳ 08-2026
11/09/2026 → kỳ 08-2026
```

Nếu yêu cầu nghiệp vụ bắt buộc **chỉ được phép chạy đúng ngày 10**, có thể thêm khóa ngày vào macro. Bản hiện tại không khóa để HR vẫn có thể chạy lại/recovery khi ngày 10 gặp lỗi mạng hoặc file chưa sẵn sàng.

---

# 13. Khuyến nghị quan trọng

API URL bạn cung cấp có thể công khai vì nó là Web App URL.

`SYNC_API_KEY` thì **không nên dùng lại key đã từng gửi trong chat hoặc nằm trong file cũ**.

Hãy tạo key mới trong Script Properties và sửa VBA.

---

## Phiên bản

**Payroll Portal – Optimized 2026 FIXED**

Các trọng điểm của bản này:

- sửa lỗi Users không được sync;
- bảo toàn mật khẩu đã đổi;
- tự tạo User lần đầu từ dữ liệu Payroll;
- giữ CCCD dạng Text;
- phục hồi CCCD 11 số → 12 số bằng `0` trong trường hợp nguồn bị mất số 0 đầu;
- loading spinner;
- không giữ thông tin đăng nhập cũ;
- hiển thị kỳ lương `Tháng MM/YYYY`;
- giữ cơ chế batch/retry;
- có macro kiểm tra file nguồn trước khi sync.
