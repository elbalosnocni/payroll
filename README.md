# PAYROLL PORTAL - FULL DEPLOYMENT

## 1. Kết quả kiểm tra file XLSB bạn gửi

File:

`PRINTING LINE 08-2026.xlsb`

Có các sheet:

- Salary
- MinWage
- PHAT (ứng)
- BANG CHAM CONG
- DAILY
- DSCNV
- Ad.Mr.Jessie
- Frontliners
- Ca đêm trước lễ
- DS7H(OLD)
- nuôi con nhỏ

File Salary có 123 cột.

Trong file thực tế:

### Salary

- D = Mã nhân viên
- CF = Họ tên
- F = Lương cơ bản
- H = Ngày làm việc
- I = Ngày lễ
- J = Nghỉ hưởng lương
- K = Nghỉ không hưởng lương
- L = OT hours
- M = Day off work
- N = Day off OT
- O = nghỉ tối thiểu vùng
- P = Day off OT night
- Q = OT night
- R = Holiday work
- S = Holiday OT
- T = Holiday OT night
- U = Night shift
- X = Tiền khác
- Y = Kỷ luật
- Z = Loyalty 2
- AA = Loyalty 5
- AB = Loyalty 10
- AC = Nhà ở
- AD = Đi lại
- AE = Chuyên cần
- AF = Phép năm còn lại / thôi việc
- AG = BHXH
- AH = BHYT
- AI = BHTN
- AJ = Khấu trừ khác
- AK = Tạm ứng
- AN = Lương tháng
- AO = OT & Night Holiday
- AQ = Quota commission
- AR = GROSS
- AT = TAX
- AW = NET B / thực lĩnh

### DSCNV

Trong file thực tế:

- B = Họ tên
- D = Mã nhân viên
- H = Số CMND/CCCD
- AF = BỘ PHẬN (ENG) / DEPARTMENT
- AG = CHỨC VỤ (ENG) / TITLE
- AH = DEPARTMENT / BỘ PHẬN (VN)
- AI = CHỨC VỤ / POSITION (VN)

Do đó code dùng:

- Department = AF
- Section = AH
- Position = AI

Mình giữ đúng tinh thần mapping của yêu cầu, nhưng đã hiệu chỉnh theo header thực tế của file bạn gửi.

---

# 2. Cấu trúc

```text
payroll_portal_full_v2/
├── gas/
│   └── Code.gs
├── web/
│   ├── index.html
│   ├── app.js
│   └── styles.css
├── vba/
│   └── SyncPayroll.bas
└── README.md
```

---

# 3. GOOGLE SHEETS

Tạo Google Spreadsheet.

Tạo sheet:

## Users

Dòng 1 chính xác:

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

Có thể đặt trên một dòng:

```text
Username | EmployeeID | FullName | PasswordHash | PasswordSalt | MustChangePassword | Active | UpdatedAt
```

## Salary

Dòng 1:

```text
PayMonth | Factory | UpdatedAt | EmployeeID | FullName | CitizenID | Department | Section | Position | TotalIncome | MonthlySalary | BasicSalary | WorkingDays | HolidayDays | PaidLeaveDays | UnpaidLeaveDays | RegionalMinimumLeaveDays | OvertimePay | OvertimeHours | RestDayHours | OvertimeRestDayHours | NightRestDayOvertimeHours | HolidayHours | HolidayOvertimeHours | NightHolidayOvertimeHours | NightShiftDays | NightOvertimeHours | OtherMoney | Discipline | Loyalty2Years | Loyalty5Years | Loyalty10Years | Housing | Transportation | AttendanceBonus | SalesCommissionBonus | SeveranceUnusedLeave | SocialInsurance | HealthInsurance | UnemploymentInsurance | PersonalIncomeTax | Advance | OtherDeductions | NetPay
```

Không cần tự nhập dữ liệu. VBA sẽ tạo.

---

# 4. GOOGLE APPS SCRIPT

Mở:

Extensions → Apps Script

Dán:

`gas/Code.gs`

Nếu Apps Script nằm trực tiếp trong Google Sheet thì không cần SPREADSHEET_ID.

Nếu là Apps Script standalone thì tạo Script Property:

```text
SPREADSHEET_ID = ID_GOOGLE_SHEET
```

---

# 5. TẠO SYNC API KEY

Trong Apps Script tạo Script Property:

```text
SYNC_API_KEY
```

Ví dụ:

```text
SYNC_API_KEY = 9b2c...một_chuỗi_ngẫu_nhiên_dài...
```

Không dùng ví dụ này thật.

Sau đó sửa VBA:

```vb
Private Const SYNC_API_KEY As String = "THAY_BANG_SYNC_API_KEY"
```

---

# 6. TẠO ADMIN

Trong Apps Script chạy hàm:

```text
setupAdmin
```

Nhập mật khẩu Admin.

Hàm sẽ tạo:

```text
ADMIN_PASSWORD_HASH
ADMIN_PASSWORD_SALT
```

trong Script Properties.

Admin password không nằm trong GitHub.

---

# 7. DEPLOY GAS

Apps Script:

Deploy → New deployment

Chọn:

```text
Web app
```

Thiết lập:

```text
Execute as: Me
Who has access: Anyone
```

Deploy.

URL phải dạng:

```text
https://script.google.com/macros/s/xxxxx/exec
```

Trong code frontend `web/app.js` đã đặt URL GAS mà bạn cung cấp.

Nếu sau này deploy một Web App mới thì thay:

```javascript
const API_URL = '...';
```

---

# 8. VBA

Mở Excel trên máy Windows có quyền truy cập:

```text
\\192.168.0.253\vn hr\
```

Alt + F11 → Insert → Module.

Import:

```text
vba/SyncPayroll.bas
```

Thay:

```vb
Private Const SYNC_API_KEY As String = "THAY_BANG_SYNC_API_KEY"
```

bằng key thật.

---

# 9. CHẠY ĐỒNG BỘ

Chạy:

```text
SyncBothFactories
```

Ví dụ:

```text
10/09/2026
```

sẽ tính:

```text
tháng trước = 08/2026
PayMonth = 08-2026
```

Tìm:

```text
\\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW\2026\SALARY 08-2026.xlsb
```

và:

```text
\\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line\2026\PRINTING LINE 08-2026.xlsb
```

Mật khẩu mở file:

```text
1234
```

---

# 10. JOIN DSCNV

Code ưu tiên dò theo:

```text
Salary.CF = DSCNV.B
```

tức:

```text
Họ tên = Họ tên
```

Nếu không tìm thấy CCCD bằng tên, code fallback theo:

```text
Salary.D = DSCNV.D
```

tức Mã nhân viên.

CCCD lấy bằng:

```vb
ws.Cells(r, 8).Text
```

để giữ số 0 đầu.

Ví dụ:

```text
058181000964
```

không bị biến thành:

```text
58181000964
```

---

# 11. MẬT KHẨU NHÂN VIÊN

Khi nhân viên xuất hiện lần đầu:

```text
Username = CCCD
Password = Mã nhân viên
```

Nhưng Google Sheet KHÔNG lưu password gốc.

Nó lưu:

```text
PasswordHash
PasswordSalt
```

Hash:

```text
SHA-256(Salt + ":" + Password)
```

và:

```text
MustChangePassword = TRUE
```

Khi đăng nhập lần đầu:

```text
CCCD + Mã nhân viên
        ↓
đăng nhập thành công
        ↓
bắt buộc đổi mật khẩu
        ↓
hash mật khẩu mới
        ↓
MustChangePassword = FALSE
```

---

# 12. ADMIN RESET

Admin đăng nhập.

Nhập CCCD.

Nhấn:

```text
Reset mật khẩu
```

Hệ thống lấy:

```text
Mã nhân viên
```

làm mật khẩu tạm thời.

Mật khẩu được hash ngay.

Sau đó:

```text
MustChangePassword = TRUE
```

Nhân viên đăng nhập bằng:

```text
CCCD
+
Mã nhân viên
```

và phải đổi mật khẩu.

---

# 13. UPDATED AT

Tất cả thời điểm cập nhật dùng:

```text
dd/MM/yyyy HH:mm:ss
```

Timezone:

```text
Asia/Ho_Chi_Minh
```

Ví dụ:

```text
10/09/2026 17:36:22
```

---

# 14. GITHUB PAGES

Upload 3 file trong thư mục `web` lên GitHub:

```text
index.html
app.js
styles.css
```

Sau đó:

Repository → Settings → Pages

chọn:

```text
Deploy from branch
main
/root
```

Website sẽ chạy trên GitHub Pages.

---

# 15. LUỒNG HỆ THỐNG

```text
                    MÁY WINDOWS
                         │
                         │ ngày 10
                         ▼
             ┌─────────────────────┐
             │      VBA Excel      │
             └──────────┬──────────┘
                        │
             ┌──────────┴──────────┐
             ▼                     ▼
        Snack XLSB           Flexible XLSB
             │                     │
             └──────────┬──────────┘
                        │
                        ▼
              Google Apps Script
                        │
                        ▼
                 Google Sheets
                  │          │
                  │          └── Users
                  └───────────── Salary
                        ▲
                        │
                        │ API
                        │
                 GitHub Pages
                        │
             ┌──────────┴──────────┐
             │                     │
          Nhân viên              Admin
             │                     │
       CCCD + password       Reset password
```

---

# 16. LƯU Ý QUAN TRỌNG

Không đặt các giá trị này lên GitHub:

```text
ADMIN_PASSWORD
SYNC_API_KEY
ADMIN_PASSWORD_HASH
ADMIN_PASSWORD_SALT
```

`SYNC_API_KEY` trong VBA vẫn có khả năng bị đọc bởi người có quyền xem VBA, vì vậy nếu hệ thống cần bảo mật cao hơn nên chuyển việc đồng bộ sang một máy chủ nội bộ hoặc cơ chế xác thực mạnh hơn.

GitHub Pages chỉ chứa giao diện.

Google Apps Script mới xử lý:

- authentication
- password hash
- session
- payslip authorization
- admin reset
- Google Sheet access
- sync API

---

# 17. ĐIỂM ĐÃ SỬA SO VỚI BẢN TRƯỚC

Bản này đã được chỉnh trực tiếp dựa trên file XLSB bạn gửi:

- Salary thực tế có 123 cột.
- CF thực sự là NAME.
- Data nhân viên thực tế bắt đầu sau phần header/group rows.
- Không còn giả định cứng rằng row 7 là dữ liệu.
- Bỏ qua các dòng TOTAL/heading không có Mã nhân viên.
- DSCNV có 165 cột.
- CCCD thực tế ở H.
- Có mã nhân viên ở D.
- Dò tên trước.
- Fallback theo mã nhân viên.
- CCCD lấy `.Text`.
- `Department / Section / Position` đã hiệu chỉnh theo header thật của file.
- Password có **salt riêng cho từng user**, tốt hơn hash SHA-256 trần.
- Sync cùng `Factory + PayMonth` sẽ thay thế kỳ cũ, không nhân đôi dữ liệu.
- Có `SyncLog`.
- Frontend escape HTML để tránh đưa dữ liệu sheet trực tiếp vào DOM không an toàn.
