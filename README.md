# Payroll GitHub Pages + Google Apps Script + VBA

Hệ thống tra cứu phiếu lương:

Excel XLSB -> VBA -> Google Apps Script -> Google Sheets -> GitHub Pages

Bản FIXED tập trung xử lý lỗi đồng bộ khiến Google Sheet không có dữ liệu:

- VBA không còn nuốt toàn bộ lỗi.
- Có log HTTP status + response GAS.
- Payload có `action: sync` rõ ràng.
- Chuẩn hóa tên để join Salary/DSCNV tốt hơn.
- Đọc CCCD dưới dạng text và GAS ép cột CCCD Plain Text.
- Không phụ thuộc ActiveSheet khi tính số cột.
- GAS kiểm tra cấu hình rõ ràng.
- GAS tạo/kiểm tra sheet và header.
- Sync có LockService tránh 2 tiến trình ghi đồng thời.
- `UpdatedAt` dùng `dd/MM/yyyy HH:mm:ss`.
- Nhân viên chỉ truy cập payroll theo CCCD trong session, không nhận CCCD/MaNV từ frontend.
- Reset password -> password = MaNV -> MustChangePassword = TRUE.
- Bắt buộc đổi password lần đầu.

## Quan trọng

ZIP không thể tự biết Google Spreadsheet ID và secret API key của bạn. Hai giá trị đó phải được đặt trong GAS Script Properties, và `SYNC_API_KEY` phải được đặt giống trong `vba/Config.bas`.

Xem `docs/SETUP_FIXED.md` trước khi chạy.
