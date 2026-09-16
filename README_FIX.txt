PAYROLL SYNC - FINAL FIX 2026-09-16

Root cause confirmed from API logs:
- VBA was manually constructing JSON numbers using locale-sensitive formatting.
- Invalid tokens such as 0. and .5 reached JSON.parse().

FINAL FIX:
1. PayrollSync.bas sends all payroll numeric fields as JSON strings (e.g. "0.5").
2. Code.gs converts those strings to real Numbers with numberValue_ before writing to Google Sheets.
3. Code.gs defensively accepts legacy .5 / -.5 / comma-decimal values too.
4. The broken VBA regex replacement pass has been removed.

DEPLOY:
- Replace Code.gs in Apps Script.
- Save.
- Deploy > Manage deployments > Edit > New version > Deploy.
- Replace/import PayrollSync.bas in the Excel workbook.
- Run RunPayrollSync.

The source Excel paths, column mapping, API URL, API key, sheet names and existing application files are retained.
