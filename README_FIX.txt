PAYROLL SYNC - FULL FIX

FILES
- Code.gs: Google Apps Script backend/API.
- index.html: GitHub Pages frontend.
- PayrollSync.bas: Excel VBA synchronization macro.

ROOT FIX FOR THE CURRENT ERROR
The VBA serializer now emits every non-ASCII UTF-16 character as JSON \uXXXX. This makes the POST body ASCII-only and avoids invalid-JSON errors caused by Excel/VBA code-page or UTF-8 conversion issues. The GAS parser also reports the JSON parser error position/message for diagnostics.

ADDITIONAL FIXES
- Removed an extra closing brace from the CSS in index.html.
- Frontend API reader now reads response text first and gives a useful error if GAS returns non-JSON.
- VBA logs the JSON payload size before POST.
- Existing API key, spreadsheet ID, sheet names, password flow, admin flow and payroll field mapping are preserved.

DEPLOYMENT
1. In the Apps Script project, replace Code.gs with the Code.gs in this package.
2. Save.
3. Deploy -> Manage deployments -> Edit the Web app deployment -> New version -> Deploy.
4. Execute as: Me.
5. Who has access: Anyone.
6. Keep using the same /exec URL in both PayrollSync.bas and index.html.
7. If the deployment URL changes, update API_URL in index.html and API_URL in PayrollSync.bas to the new /exec URL.
8. Run setup once if the system has not been initialized.

VBA
1. Open the Excel workbook that contains the macro.
2. Replace/import PayrollSync.bas.
3. Run RunPayrollSync.
4. Check PayrollSync.log next to the workbook if an error occurs. The log now includes JSON size and the full API response prefix.

IMPORTANT
After changing Code.gs, Apps Script must be redeployed as a NEW VERSION. Editing/saving Code.gs alone does not change an existing deployed Web App version.
