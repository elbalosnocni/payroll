PAYROLL SYNC VBA v2.1 - HEADER MAPPING
======================================

Muc tieu
--------
PayrollSync.bas khong con Private Const COL_... hoac DSCNV_*_COL.
VBA tu tim cot theo Header trong 6 dong dau cua sheet (SHEET_START_ROW = 7).

Tuong thich file Excel hien tai
--------------------------------
Neu khong tim thay Header, VBA tu dong dung vi tri legacy cua he thong cu va ghi
CANH BAO vao PayrollSync.log. Vi vay file Excel hien tai van co the chay ngay.
Khong can sua VBA chi vi chen them / di chuyen cot.

Header alias
------------
VBA nhan nhieu ten cho cung mot field, vi du:
Employee Code / EmployeeCode / Emp Code / Ma nhan vien / Ma NV
Basic Salary / BasicSalary / Luong co ban
Working Days / WorkingDays / Ngay cong
...

Neu muon doi ten Header
-----------------------
Chi can them alias vao ham SalaryFieldSpecs_ trong PayrollSync.bas.
Khong can khai bao so cot.

Cot bat buoc
------------
EmployeeCode va FullName la 2 field bat buoc. Neu khong tim thay va fallback
legacy cung khong hop le, VBA dung voi loi ro rang.

Cot tuy chon
------------
Cac field luong/OT/phu cap/khau tru la optional. Neu khong tim thay Header,
VBA dung 0 (hoac fallback legacy neu file cu dang dung cot do).

DSCNV
-----
Name, Citizen ID, Department, Section, Position cung duoc map theo Header.
Neu khong tim thay, VBA fallback ve cot legacy 2/8/32/33/34.

Log
---
Moi lan chay VBA ghi mapping vao PayrollSync.log, vi du:
MAP basicSalary => col 7 | Header='Basic Salary'
CANH BAO: Khong tim thay Header ... Dung cot legacy ...

Them cot moi
------------
Neu cot moi chua phai field he thong, VBA se nhan dien Header va dua vao
extraData trong JSON. Backend phai co PayrollConfig/ExtraData support thi moi
luu va hien thi cot nay trong Google Sheet/phieu luong.

Luu y
-----
- API_URL, SYNC_API_KEY, FILE_PASSWORD, ROOT_PATH van giu nguyen de tuong thich.
- Không cần tham chiếu thư viện VBA ngoài; code dùng late binding.
- Nên import file PayrollSync.bas mới thay cho module cũ sau khi backup.
