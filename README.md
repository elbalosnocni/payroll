# Hệ thống tra cứu phiếu lương (GitHub Pages + Google Apps Script + VBA)

Hệ thống cho phép nhân viên đăng nhập bằng **số Căn cước công dân** để xem
phiếu lương của chính mình, dữ liệu được đồng bộ tự động hằng tháng từ file
Excel (`.xlsb`) qua VBA, trung chuyển qua Google Sheets + Google Apps
Script, hiển thị trên trang tĩnh GitHub Pages.

```
Excel (.xlsb, 2 xưởng) --[VBA, ngày 10 hằng tháng]--> Google Apps Script --> Google Sheets
                                                                                   │
                                                                                   ▼
                                                            GitHub Pages (đăng nhập CCCD) 
```

## Cấu trúc thư mục

```
payroll_github_gas/
├── README.md
├── gas/            Google Apps Script (API backend)
│   ├── Code.gs         Router doGet/doPost
│   ├── Config.gs       Cấu hình chung
│   ├── Auth.gs         Đăng nhập, đổi mật khẩu, session
│   ├── Payroll.gs      Lấy phiếu lương nhân viên
│   ├── Admin.gs        Chức năng Admin
│   ├── Sync.gs         Nhận dữ liệu đẩy lên từ VBA
│   └── Utils.gs        Hàm tiện ích dùng chung
├── web/            Frontend tĩnh (deploy lên GitHub Pages)
│   ├── index.html
│   ├── app.js          Giao diện xem phiếu lương
│   ├── auth.js         Đăng nhập / API client / session
│   ├── admin.js        Giao diện quản trị
│   └── styles.css
├── vba/            Macro đồng bộ dữ liệu từ Excel
│   ├── SyncPayroll.bas
│   ├── Config.bas
│   └── Json.bas
└── docs/
    ├── GOOGLE_SHEET_SETUP.md
    ├── GAS_DEPLOY.md
    ├── GITHUB_PAGES.md
    └── WINDOWS_TASK_SCHEDULER.md
```

## Thứ tự triển khai

1. **Google Sheet** — làm theo `docs/GOOGLE_SHEET_SETUP.md`: tạo Spreadsheet,
   dán code `gas/`, chạy `setupSheets`, tạo tài khoản Admin đầu tiên.
2. **Deploy Apps Script** — làm theo `docs/GAS_DEPLOY.md`: deploy Web App,
   lấy URL, dán vào `web/auth.js` và `vba/Config.bas`.
3. **GitHub Pages** — làm theo `docs/GITHUB_PAGES.md`: đẩy thư mục `web/`
   lên GitHub, bật Pages.
4. **VBA + Task Scheduler** — làm theo `docs/WINDOWS_TASK_SCHEDULER.md`: tạo
   file `.xlsm` chứa macro, cấu hình `vba/Config.bas`, lên lịch chạy ngày 10
   hằng tháng.

## Luồng hoạt động

- **Ngày 10 hằng tháng**: VBA tự xác định tháng liền trước tháng hiện tại,
  tìm 2 file `SALARY <MM-YYYY>.xlsb` (xưởng Snack) và
  `PRINTING LINE <MM-YYYY>.xlsb` (xưởng Flexible) trong thư mục năm tương
  ứng, mở bằng mật khẩu cấu hình sẵn, đọc 2 sheet `Salary` và `DSCNV` từ
  dòng 7, ghép theo Họ tên để lấy CCCD/phòng ban/bộ phận/chức vụ, rồi POST
  toàn bộ dữ liệu (JSON) lên Google Apps Script.
- **Apps Script** ghi dữ liệu vào Google Sheets (`Employees`, `Payroll`),
  tự tạo tài khoản mới với mật khẩu ban đầu = Mã nhân viên nếu chưa tồn tại,
  giữ nguyên mật khẩu nếu nhân viên đã đổi trước đó.
- **Nhân viên** vào trang GitHub Pages, đăng nhập bằng CCCD + mật khẩu. Lần
  đầu (hoặc sau khi Admin reset), hệ thống bắt buộc đổi mật khẩu trước khi
  xem được phiếu lương. Server luôn xác định nhân viên qua token phiên đăng
  nhập, không tin dữ liệu CCCD/MaNV do trình duyệt gửi lên khi truy vấn dữ
  liệu lương — nhân viên A không thể sửa request để xem lương nhân viên B.
- **Admin** đăng nhập cùng cổng, thấy thêm mục quản trị: danh sách/tìm nhân
  viên, reset mật khẩu (về Mã nhân viên, bắt buộc đổi lại), xem trạng thái
  đồng bộ, xem Audit log (đăng nhập, đổi mật khẩu, reset, đồng bộ...).

## Lưu ý bảo mật quan trọng

- Mật khẩu nhân viên được băm (SHA-256 lặp nhiều vòng + salt riêng + pepper
  chung), không lưu plaintext.
- Endpoint `/sync` (VBA → GAS) được bảo vệ bằng `SYNC_API_KEY` riêng, khác
  với mật khẩu file Excel (`1234`) và khác cơ chế đăng nhập nhân viên.
- Session dùng token ngẫu nhiên lưu trong `CacheService` (phía server, hết
  hạn sau 8 giờ), không lưu trực tiếp trong Sheet.
- File `.xlsb` gốc và mật khẩu mở file (`1234`) không bao giờ được gửi ra
  ngoài mạng nội bộ — chỉ dữ liệu đã trích xuất (JSON) được gửi lên Google.
- Trước khi đưa vào dùng thật: đổi `SYNC_API_KEY`, `PASSWORD_PEPPER` trong
  Script Properties thành các chuỗi ngẫu nhiên đủ dài, và cân nhắc đổi mật
  khẩu mở file `.xlsb` (`1234`) sang mật khẩu mạnh hơn nếu có thể.

## Những phần cần bạn kiểm tra/điều chỉnh lại thực tế

Vì không có quyền truy cập trực tiếp vào file `.xlsb` và Google Sheet thật
của bạn, code được viết dựa hoàn toàn trên mô tả cột/sheet đã cung cấp. Khi
triển khai thật, nên:

- Chạy thử `SyncPayroll.RunSync` với 1 file mẫu, đối chiếu số liệu JSON gửi
  lên (xem `%TEMP%\SyncPayroll.log`) với phiếu lương thật để xác nhận đúng
  cột.
- Kiểm tra định dạng cột CCCD trong DSCNV — nếu cách lưu số 0 đầu khác với
  mô tả (dùng `Format(Text)` hay dấu `'`), có thể cần chỉnh `GetCellText`
  trong `SyncPayroll.bas`.
- Nếu 1 Họ tên trùng nhau giữa nhiều nhân viên trong DSCNV, việc ghép theo
  tên (thay vì mã nhân viên) có thể bị sai — cân nhắc bổ sung cột Mã NV
  trong sheet DSCNV để ghép theo mã thay vì tên nếu có thể.
