PAYROLL WEB APP V2 - INSTALL / MIGRATION GUIDE
===============================================

FILES
-----
1. Code.gs                 -> Google Apps Script backend v2
2. index.html              -> GitHub Pages frontend v2
3. PayrollSync.bas         -> VBA sync v2
4. RunPayrollSync.vbs      -> giữ nguyên launcher
5. PayrollCustomFields.csv -> mẫu cấu hình cột Excel tùy chỉnh

WHAT V2 ADDS
------------
- Session 60 minutes + CacheService acceleration.
- Employee failed-login counter and temporary lock after 5 failed attempts / 15 minutes.
- Admin failed-login counter and temporary lock.
- Admin roles: SUPER_ADMIN, HR_ADMIN, PAYROLL_ADMIN, AUDITOR.
- Admin Dashboard.
- Employee lock/unlock.
- Admin create / lock / unlock.
- Admin change own password.
- Audit logging for security actions.
- PayrollConfig sheet created automatically.
- SystemSettings sheet created automatically.
- Payroll ExtraData column for custom payroll fields.
- Existing Sheets are migrated by APPENDING missing headers; old data is not deleted by migration.
- index.html and PayrollSync.bas use the same current Web App URL from the original index.html.

DEPLOYMENT
----------
A. Google Apps Script
1. Open the existing Apps Script project.
2. Replace Code.gs with the v2 Code.gs from this package.
3. Save.
4. Deploy -> Manage deployments -> Edit the Web app deployment.
5. Execute as: Me.
6. Who has access: Anyone who needs to use the internal system (normally Anyone with the link, subject to your company policy).
7. Deploy a NEW VERSION of the same Web App deployment.
8. Do not create a second deployment unless necessary. Keeping one endpoint avoids frontend/VBA mismatch.

B. First initialization
- Open the Web App once.
- The backend automatically creates/migrates:
  Employees, Payroll, Admins, Sessions, AuditLogs, SyncStatus, PayrollConfig, SystemSettings.
- Existing rows are preserved.
- The original admin account remains compatible. If it was already created, its password is NOT overwritten.

C. GitHub Pages
- Replace index.html with v2 index.html.
- Confirm API_URL at the top of index.html is the same URL as the deployed Apps Script Web App.

D. VBA
- Import/replace PayrollSync.bas.
- The macro keeps the same RunPayrollSync entry point.
- RunPayrollSync.vbs can continue to call the same macro.

ADDING A NEW PAYROLL FIELD - V2
--------------------------------
For a new Excel field that should be sent to the payslip, you no longer edit Code.gs or index.html.

1. In the workbook containing the VBA module, create a sheet named:
   PayrollCustomFields

2. Put this header row:
   FieldKey | ExcelColumn | Label | Type | Active

3. Example:
   PhoneAllowance | 50 | Phụ cấp điện thoại | money | TRUE
   MealAllowance  | 51 | Phụ cấp ăn ca      | money | TRUE
   Note           | 52 | Ghi chú             | text  | TRUE

4. ExcelColumn is the actual column number in the Salary sheet.
5. Type supports: money, number, text.
6. Run Payroll Sync.
7. The values are stored in Payroll!ExtraData.
8. The employee payslip automatically displays them under D. Thông tin bổ sung.
9. If you want a friendly label/type/group in Admin configuration, add a matching row to PayrollConfig.

IMPORTANT
---------
- The Salary workbook column number is 1-based Excel column number.
- FieldKey should contain only letters/numbers/underscore where possible.
- Do not reuse a standard payroll field name unless you intentionally want to override it.
- Existing standard fields continue to work exactly as before.

SECURITY
--------
The old VBA sync API key is retained as a fallback for compatibility. After confirming v2 works, move to a new secret:
1. In Apps Script, set SystemSettings row:
   SYNC_API_KEY = <new long random secret>
2. Change SYNC_API_KEY in PayrollSync.bas to the same value.
3. Do not publish PayrollSync.bas publicly.
4. Change the default Admin password immediately.

TEST CHECKLIST
--------------
1. Employee login.
2. Employee password change.
3. View current payslip.
4. Switch historical months.
5. Admin login.
6. Dashboard.
7. Employee search.
8. Lock/unlock employee.
9. Reset employee password.
10. Create a non-SUPER_ADMIN account.
11. Admin password change.
12. View Audit Log.
13. Run VBA sync for one month.
14. Verify Payroll row count and ExtraData.
15. Verify a custom field appears under D. Thông tin bổ sung.

ROLLBACK
--------
Keep the original ZIP as a backup. If rollback is required, restore the original Code.gs/index.html/PayrollSync.bas and redeploy the previous Apps Script version.
