# Google Sheet Setup

Chạy `setupSheets()`.

### Employees
`CCCD | MaNV | HoTen | Xuong | PhongBan | BoPhan | ChucVu | PasswordHash | PasswordSalt | MustChangePassword | Role | UpdatedAt`

Định dạng cột CCCD là Plain text. Code cũng ghi CCCD dưới dạng chuỗi.

### Payroll
Một dòng cho mỗi `MaNV + Thang`.

### Admin
Thêm dòng admin với Role=`ADMIN`, MaNV ví dụ `ADMIN01`, CCCD dạng text. Chạy `setInitialAdmin("ADMIN01")`. Mật khẩu ban đầu = ADMIN01 và bắt buộc đổi.

Không nhập PasswordHash/Salt bằng tay.
