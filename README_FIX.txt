PAYROLL SYNC - MONTHLY ARCHIVE / NO DUPLICATE PAYSLIPS

FIXES:
1. Re-syncing the same salary month replaces the entire month instead of appending duplicates.
2. Existing Google Sheets date-formatted SalaryMonth cells are normalized (e.g. 2026-07-31T17:00:00.000Z -> 08-2026).
3. SalaryMonth is stored as TEXT so Google Sheets will not convert 08-2026 to a date.
4. Incoming duplicate EmployeeCode rows are collapsed to one payslip per employee per month.
5. Historical payslips are retained by SalaryMonth in the Payroll sheet.
6. Employee portal now shows a month selector so employees can view previous months.
7. Existing latest-payslip behavior is preserved: the newest month is selected by default.

IMPORTANT ONE-TIME CLEANUP:
- Deploy the new Code.gs as a NEW VERSION of the Apps Script Web App.
- Do NOT manually delete August data first.
- Run the VBA payroll sync for 08-2026 once.
- The server will remove ALL existing 08-2026 Payroll rows (including old duplicates/date-formatted rows) and write the fresh data once.
- Older months such as 07-2026 remain untouched and can still be viewed.

FILES:
- Code.gs: backend/API
- index.html: employee web portal with monthly payslip history
- PayrollSync.bas: VBA sync module (unchanged in this package because duplicate prevention is enforced server-side)
