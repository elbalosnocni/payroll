# Payroll Slip - Optimized v2.3

## Bộ mã này hợp nhất từ
- Gihub_Slip.zip
- Gihub_Slip_v2.2.zip

## Thay đổi chính
1. Giữ Header Mapping + ExtraData của v2.2.
2. Sửa lỗi VBA object assignment:
   `Set specs = SalaryFieldSpecs_()`.
3. Sửa logic xác định dòng dữ liệu Salary:
   - Không còn dừng ở TOTAL đầu tiên.
   - Cho phép nhiều TOTAL/SUBTOTAL trong cùng sheet.
   - Chỉ coi dòng có Employee Code + Full Name là dòng nhân viên.
4. DSCNV ưu tiên map theo Employee Code; nếu không có thì fallback theo tên.
5. Không format toàn bộ số hàng của Payroll mỗi lần sync; chỉ format vùng dữ liệu vừa ghi.
6. Giữ PayrollConfig để frontend tự nhận các cột Excel mới trong ExtraData.
7. Giữ cơ chế xóa/thay thế toàn bộ payroll của tháng trước khi ghi lại, tránh duplicate.
8. Thêm giới hạn 10.000 record cho API sync.

## Thứ tự triển khai
1. Code.gs -> Apps Script.
2. index.html -> GitHub Pages.
3. PayrollSync.bas -> file PayrollSync.xlsm.
4. Kiểm tra `SYNC_API_KEY`, `SETUP_KEY`, `SPREADSHEET_ID` theo hệ thống của bạn.
5. Chạy setup một lần.
6. Chạy VBA sync thử với một tháng dữ liệu mẫu.

## Lưu ý bảo mật
Sau khi triển khai, nên thay API key / setup key / mật khẩu admin mặc định trong Apps Script và VBA.
