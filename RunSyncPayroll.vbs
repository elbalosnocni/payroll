Option Explicit
Dim xl, wb
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then WScript.Quit 1
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = True
xl.AskToUpdateLinks = False
Set wb = xl.Workbooks.Open("X:\DS + PN + TP - 2014\AnnualLeaveSync.xlsm", False, False)
If Err.Number <> 0 Then xl.Quit : WScript.Quit 2
Err.Clear
xl.Run "'" & wb.Name & "'!SyncAnnualLeaveToGoogle"
If Err.Number <> 0 Then wb.Close False : xl.Quit : WScript.Quit 3
wb.Save
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
WScript.Quit 0
