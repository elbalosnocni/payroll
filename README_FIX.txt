PAYROLL SYNC - FIX JSON

1. Replace Code.gs in Apps Script.
2. Deploy -> Manage deployments -> Edit -> New version -> Deploy.
3. Execute as: Me; Who has access: Anyone.
4. Replace PayrollSync.bas in VBA.
5. Run RunPayrollSync.

This version fixes:
- VBA UTF-8 JSON POST using ADODB.Stream bytes (BOM removed).
- JSON escaping for all control characters.
- Content-Type application/json; charset=utf-8.
- GAS parser strips BOM/whitespace and reports JSON prefix on failure.
- API key matches between VBA and GAS.
- CreateObject line syntax fixed.

Check PayrollSync.log in the same folder as the Excel workbook if an error occurs.
