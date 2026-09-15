# Windows Task Scheduler

Tạo `C:\PayrollSync\RunSync.vbs`:

```vbscript
Dim xl, wb
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open("C:\PayrollSync\PayrollSyncRunner.xlsm")
xl.Run "SyncPayroll.RunScheduledSync"
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
```

Task:
- Monthly
- Day 10
- 08:00 hoặc sau khi file lương đã hoàn tất
- Run whether user is logged on or not
- Account Windows phải có quyền đọc `\\192.168.0.253\vn hr\...`

`RunScheduledSync` có guard ngày 10; `RunSync` vẫn cho phép chạy thử thủ công.
