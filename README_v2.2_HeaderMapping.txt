PAYROLL WEB APP v2.2 - DYNAMIC HEADER MAPPING
Date: 2026-10-02

1. THAY FILE
- Code.gs: thay toàn bộ Code.gs trên Apps Script.
- index.html: thay toàn bộ index.html trên GitHub Pages.
- PayrollSync.bas: dùng bản Header Mapping v2.1 đã có trong bộ trước.

2. NGUYÊN TẮC
VBA không phụ thuộc vị trí cột Excel cho các field chuẩn. VBA tìm theo Header/Alias.
Các Header không thuộc field chuẩn được đưa vào JSON extraData.

GAS lưu extraData vào cột Payroll!ExtraData dưới dạng JSON.
GAS tự tạo sheet PayrollConfig:
FieldKey | Label | Type | Active | Order | CreatedAt | UpdatedAt

3. THÊM CỘT EXCEL
Ví dụ thêm Header: Phone Allowance
- Không sửa Code.gs.
- Không sửa index.html.
- Không sửa số cột trong VBA.
- VBA tự đưa vào extraData với key phone_allowance.
- Sau lần Sync, GAS tự thêm key vào PayrollConfig.
- Phiếu lương tự hiển thị trong nhóm D. Các khoản bổ sung từ Excel.

4. ĐỔI VỊ TRÍ CỘT
Có thể chèn/xóa/di chuyển các cột Excel. VBA vẫn tìm theo Header.
Legacy column chỉ là fallback cho các field chuẩn nếu Header không tìm thấy.

5. PAYROLLCONFIG
Có thể sửa Label để đổi tên hiển thị trên phiếu lương.
Type:
- money
- number
- text
- boolean
- auto
Active FALSE để ẩn field khỏi schema (dữ liệu vẫn lưu trong ExtraData).
Order dùng để sắp xếp field.

6. MIGRATION
Code.gs không ghi đè header cũ. Nếu thiếu ExtraData hoặc PayrollConfig, hệ thống tự thêm.
Nên backup Google Sheet trước khi deploy.

7. TRIỂN KHAI
- Apps Script: paste Code.gs -> Save -> Deploy > Manage deployments -> Edit -> New version.
- GitHub Pages: thay index.html.
- VBA: tiếp tục dùng PayrollSync Header Mapping v2.1.
- Sync thử 1 tháng test trước khi chạy production.


8. BAN SUA 02/10/2026
- Sua loi VBA "Too many line continuations": SalaryFieldSpecs_ khong con dung mot Array() dai 38 dong. Mapping duoc tao bang Collection + SalarySpecAdd_.
- Bao toan alias co dau va khong dau de nhan dien Header Excel.
- GrossIncome van map dung cot AR (44) khi khong tim thay Header.
- Them FindPayrollDataLastRow_: doc tu dong 7 va dung ngay truoc dong co cot B bat dau bang TOTAL, TOTAL1, TOTAL2, TOTAL3. Neu khong co dong TOTAL thi fallback ve dong cuoi cua cot FullName.
- Rut gon tinh OtherIncome thanh ham OtherIncomeTotal_, van giu dung tong cac nhom muc III.
- Giao dien phieu luong duoc giu nguyen: A noi bat, B noi bat, C - Luong thuc linh noi bat nhat; tien cua I/II/III hien rieng ben phai tieu de.
- Code.gs van tinh OtherIncome = tong cac subgroup muc III, khong phu thuoc cot OtherIncome trong Excel.
