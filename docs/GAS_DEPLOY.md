# GAS Deploy

1. Script Properties:
   - SPREADSHEET_ID
   - SYNC_API_KEY
   - PASSWORD_PEPPER
2. Run `setupSheets`.
3. Deploy → New deployment → Web app.
4. Execute as: Me.
5. Access: Anyone.
6. Sau mỗi lần sửa code: Deploy → Manage deployments → Edit → New version → Deploy.

Frontend dùng POST với `Content-Type: text/plain;charset=utf-8` để tránh preflight OPTIONS của Apps Script.
