Option Explicit

Dim excelApp
Dim workbook
Dim workbookPath

workbookPath = "D:\vtc\github\PayrollSync.xlsm"

On Error Resume Next

Set excelApp = CreateObject("Excel.Application")

If Err.Number <> 0 Then
    WScript.Echo "Khong the khoi dong Microsoft Excel."
    WScript.Quit 1
End If

excelApp.Visible = False
excelApp.DisplayAlerts = False
excelApp.EnableEvents = False
excelApp.ScreenUpdating = False

Set workbook = excelApp.Workbooks.Open( _
    workbookPath, _
    0, _
    False)

If Err.Number <> 0 Then
    WScript.Echo "Khong the mo file: " & workbookPath
    excelApp.Quit
    WScript.Quit 1
End If

Err.Clear

excelApp.Run "RunPayrollSync"

If Err.Number <> 0 Then
    WScript.Echo "Macro dong bo gap loi: " & Err.Description
    Err.Clear
End If

workbook.Close False

excelApp.Quit

Set workbook = Nothing
Set excelApp = Nothing

WScript.Quit 0
