# Payroll Portal - Optimized

## Files
- `gas/Code.gs` — Google Apps Script API
- `web/index.html` — GitHub Pages UI
- `web/app.js` — frontend logic
- `web/styles.css` — UI styling
- `vba/SyncPayroll.bas` — Excel VBA sync from XLSB

## Main fixes
1. VBA reads Salary and DSCNV into arrays instead of cell-by-cell loops.
2. DSCNV mapping is corrected:
   - B = Họ tên
   - H = CCCD
   - AF = Phòng ban
   - AG = Bộ phận
   - AH = Chức vụ
3. Join uses EmployeeID first, then normalized name as fallback.
4. XLSB password is `1234`.
5. Payroll is sent in batches of 500.
6. GAS clears only the target Factory + PayMonth on batch 1.
7. Later batches UPSERT by Factory + PayMonth + EmployeeID, so one batch no longer deletes previous batches.
8. VBA retries failed API requests up to 3 times.
9. GAS uses LockService during each sync request.
10. Server blocks payslip access until the employee changes the initial/reset password.
11. Citizen ID is written as text in Salary.
12. SyncLog records batch details.
13. Frontend handles session expiration and API timeout.

## IMPORTANT API KEY
The VBA file intentionally contains:
`PUT_YOUR_RANDOM_SYNC_API_KEY_HERE`

Create a NEW random API key and put exactly the same value in:
Google Apps Script -> Project Settings -> Script Properties:
`SYNC_API_KEY`

Do not use or publish the old key.

## GAS setup
1. Open the Apps Script project.
2. Replace Code.gs with `gas/Code.gs`.
3. Set Script Properties:
   - `SPREADSHEET_ID` = Google Sheet ID (optional if bound script)
   - `SYNC_API_KEY` = your new random secret
4. Run `setAdminPassword("your-admin-password")` once manually.
5. Deploy as Web app:
   - Execute as: Me
   - Who has access: Anyone
6. Confirm the deployed URL matches the URL in `web/app.js` and `vba/SyncPayroll.bas`.

## Google Sheets headers
Code.gs automatically creates/repairs:
- Users
- Salary
- SyncLog

Do not manually reorder Salary columns after deployment.

## GitHub Pages
Upload:
- index.html
- app.js
- styles.css

to the same folder.

## Monthly sync
`SyncBothFactories` automatically uses the previous calendar month.
Example:
10/09/2026 -> 08-2026.

VBA still requires Excel/Windows to be running. To make it automatic on the 10th, use Windows Task Scheduler to open the workbook and run `SyncBothFactories`.

## Security note
The API key is embedded in VBA and therefore cannot be treated as a secret once distributed to users who can inspect the VBA project. It is still useful as a sync gate, but for stronger security move the sync process to a controlled server/service or protect the VBA project and rotate the key periodically.
