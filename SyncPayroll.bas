Option Explicit

' ============================================================
' PAYROLL SYNC - VBA / Excel XLSB
' Version: Fully Optimized Batch Sync (Memory & Connection Efficient)
' ============================================================

Private Const GAS_URL As String = _
"https://google.com"

' ⚠️ Hãy thay chuỗi này bằng mã bảo mật của bạn
Private Const SYNC_API_KEY As String = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE"
Private Const XLS_PASSWORD As String = "1234"
Private Const ROOT As String = "\\192.168.0.253\vn hr\"

Private Const BATCH_SIZE As Long = 500
Private Const MAX_RETRY As Long = 3

' Salary columns
Private Const COL_EMPLOYEE_ID As Long = 4
Private Const COL_BASIC As Long = 6
Private Const COL_WORKING_DAYS As Long = 8
Private Const COL_HOLIDAY_DAYS As Long = 9
Private Const COL_PAID_LEAVE As Long = 10
Private Const COL_UNPAID_LEAVE As Long = 11
Private Const COL_OT_HOURS As Long = 12
Private Const COL_REST_HOURS As Long = 13
Private Const COL_OT_REST_HOURS As Long = 14
Private Const COL_MIN_WAGE_LEAVE As Long = 15
Private Const COL_NIGHT_REST_OT As Long = 16
Private Const COL_NIGHT_OT As Long = 17
Private Const COL_HOLIDAY_HOURS As Long = 18
Private Const COL_HOLIDAY_OT As Long = 19
Private Const COL_NIGHT_HOLIDAY_OT As Long = 20
Private Const COL_NIGHT_DAYS As Long = 21
Private Const COL_OTHER_MONEY As Long = 24
Private Const COL_DISCIPLINE As Long = 25
Private Const COL_LOYALTY2 As Long = 26
Private Const COL_LOYALTY5 As Long = 27
Private Const COL_LOYALTY10 As Long = 28
Private Const COL_HOUSING As Long = 29
Private Const COL_TRANSPORT As Long = 30
Private Const COL_ATTENDANCE As Long = 31
Private Const COL_SEVERANCE As Long = 32
Private Const COL_SOCIAL As Long = 33
Private Const COL_HEALTH As Long = 34
Private Const COL_UNEMPLOYMENT As Long = 35
Private Const COL_OTHER_DEDUCT As Long = 36
Private Const COL_ADVANCE As Long = 37
Private Const COL_MONTHLY_SALARY As Long = 40
Private Const COL_OT_PAY As Long = 41
Private Const COL_COMMISSION As Long = 43
Private Const COL_TOTAL_INCOME As Long = 44
Private Const COL_TAX As Long = 46
Private Const COL_NET_PAY As Long = 49
Private Const COL_NAME As Long = 84

' DSCNV columns
Private Const DS_NAME As Long = 2
Private Const DS_CCCD As Long = 8
Private Const DS_DEPARTMENT As Long = 32 ' AF
Private Const DS_SECTION As Long = 33    ' AG
Private Const DS_POSITION As Long = 34   ' AH
Private Const DS_EMPLOYEE_ID As Long = 4 ' D

' Module-level lookup maps & HTTP object reused for performance
Private gDSByID As Object
Private gDSByName As Object
Private m_Http As Object 

' ============================================================
' ENTRY POINTS
' ============================================================

Public Sub SyncBothFactories()
    Dim payMonth As String: payMonth = PreviousPayrollMonth_()
    Dim y As String: y = Right$(payMonth, 4)

    Dim snackPath As String: snackPath = ROOT & "SALARY - 2014 - 2015\VNLWW\" & y & "\SALARY " & payMonth & ".xlsb"
    Dim flexPath As String: flexPath = ROOT & "SALARY - 2014 - 2015\Printing line\" & y & "\PRINTING LINE " & payMonth & ".xlsb"

    Dim msg As String: msg = "KỲ LƯƠNG: " & payMonth & vbCrLf & vbCrLf
    Dim okSnack As Boolean, okFlex As Boolean

    ' Khởi tạo HTTP Object dùng chung một lần duy nhất cho toàn bộ phiên làm việc
    Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")

    If FileExists_(snackPath) Then
        okSnack = SyncFactoryFile(snackPath, "Snack", payMonth)
        msg = msg & IIf(okSnack, "✓ Snack: thành công", "✗ Snack: thất bại") & vbCrLf
    Else
        msg = msg & "✗ Snack: không tìm thấy file" & vbCrLf
    End If

    If FileExists_(flexPath) Then
        okFlex = SyncFactoryFile(flexPath, "Flexible", payMonth)
        msg = msg & IIf(okFlex, "✓ Flexible: thành công", "✗ Flexible: thất bại") & vbCrLf
    Else
        msg = msg & "✗ Flexible: không tìm thấy file" & vbCrLf
    End If

    Set m_Http = Nothing ' Giải phóng bộ nhớ mạng
    Application.StatusBar = False
    MsgBox msg, IIf(okSnack Or okFlex, vbInformation, vbCritical), "Payroll Sync Complete"
End Sub

Public Sub SyncSnackOnly()
    Dim pm As String: pm = PreviousPayrollMonth_()
    Dim p As String: p = ROOT & "SALARY - 2014 - 2015\VNLWW\" & Right$(pm, 4) & "\SALARY " & pm & ".xlsb"
    If Not FileExists_(p) Then MsgBox "Không tìm thấy file:" & vbCrLf & p, vbCritical: Exit Sub
    
    Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call SyncFactoryFile(p, "Snack", pm)
    Set m_Http = Nothing
End Sub

Public Sub SyncFlexibleOnly()
    Dim pm As String: pm = PreviousPayrollMonth_()
    Dim p As String: p = ROOT & "SALARY - 2014 - 2015\Printing line\" & Right$(pm, 4) & "\PRINTING LINE " & pm & ".xlsb"
    If Not FileExists_(p) Then MsgBox "Không tìm thấy file:" & vbCrLf & p, vbCritical: Exit Sub
    
    Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call SyncFactoryFile(p, "Flexible", pm)
    Set m_Http = Nothing
End Sub

' ============================================================
' CORE SYNC ENGINE
' ============================================================

Public Function SyncFactoryFile( _
    ByVal filePath As String, _
    ByVal factory As String, _
    ByVal payMonth As String) As Boolean

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        Password:=XLS_PASSWORD, _
        ReadOnly:=True, _
        UpdateLinks:=0, _
        IgnoreReadOnlyRecommended:=True)

    Dim wsSalary As Worksheet: Set wsSalary = wb.Worksheets("Salary")
    Dim wsDS As Worksheet: Set wsDS = wb.Worksheets("DSCNV")

    Set gDSByID = CreateObject("Scripting.Dictionary")
    Set gDSByName = CreateObject("Scripting.Dictionary")
    gDSByID.CompareMode = vbTextCompare
    gDSByName.CompareMode = vbTextCompare

    BuildDSMaps_ wsDS

    Dim lastRow As Long
    lastRow = wsSalary.Cells(wsSalary.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row

    If lastRow < 1 Then
        MsgBox factory & ": không có dữ liệu Salary.", vbExclamation
        GoTo FAIL_EXIT
    End If

    ' Tối ưu: Đọc toàn mảng một lần duy nhất
    Dim salaryData As Variant
    salaryData = wsSalary.Range(wsSalary.Cells(1, 1), wsSalary.Cells(lastRow, COL_NAME)).Value2

    Dim totalRows As Long
    totalRows = CountSalaryRows_(salaryData, lastRow)

    If totalRows = 0 Then
        MsgBox factory & ": không có nhân viên hợp lệ.", vbExclamation
        GoTo FAIL_EXIT
    End If

    Dim rowNums() As Long: ReDim rowNums(1 To totalRows)
    Dim i As Long, r As Long
    Dim emp As String, nm As String

    i = 0
    For r = 1 To lastRow
        emp = CleanText_(CStrSafe_(salaryData(r, COL_EMPLOYEE_ID)))
        nm = CleanText_(CStrSafe_(salaryData(r, COL_NAME)))
        If emp <> "" And nm <> "" Then
            i = i + 1
            rowNums(i) = r
        End If
    Next r

    Dim totalBatches As Long
    totalBatches = (totalRows + BATCH_SIZE - 1) \ BATCH_SIZE

    Dim batchId As String
    batchId = factory & "_" & Replace(payMonth, "-", "") & "_" & Format$(Now, "yyyymmddhhnnss")

    Dim startIndex As Long, endIndex As Long, batchNo As Long
    batchNo = 0

    For startIndex = 1 To totalRows Step BATCH_SIZE
        batchNo = batchNo + 1
        endIndex = startIndex + BATCH_SIZE - 1
        If endIndex > totalRows Then endIndex = totalRows

        ' Tối ưu: Sử dụng cơ chế nối mảng hiệu năng cao cho chuỗi lô dữ liệu lớn
        Dim payload As String
        payload = BuildBatchPayload_(salaryData, rowNums, startIndex, endIndex, factory, payMonth, batchId, batchNo, totalBatches)

        Dim response As String
        If Not PostWithRetry_(payload, response) Then
            MsgBox factory & " thất bại tại lô " & batchNo & "/" & totalBatches & vbCrLf & response, vbCritical
            GoTo FAIL_EXIT
        End If

        If Not ResponseOK_(response) Then
            MsgBox factory & " máy chủ từ chối lô " & batchNo & "/" & totalBatches & vbCrLf & response, vbCritical
            GoTo FAIL_EXIT
        End If

        Application.StatusBar = factory & " - Đang tải: lô " & batchNo & "/" & totalBatches & " (" & endIndex & "/" & totalRows & " nhân viên)"
    Next startIndex

    SyncFactoryFile = True
    GoTo CLEAN_EXIT

FAIL_EXIT:
    SyncFactoryFile = False

CLEAN_EXIT:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set gDSByID = Nothing
    Set gDSByName = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Function

EH:
    SyncFactoryFile = False
    MsgBox "Lỗi SyncFactoryFile tại " & factory & ":" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    Resume CLEAN_EXIT
End Function

' ============================================================
' HELPER METHODS & OPTIMIZATIONS
' ============================================================

Private Sub BuildDSMaps_(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, DS_EMPLOYEE_ID).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim data As Variant
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, DS_POSITION)).Value2

    Dim r As Long, employeeId As String, fullName As String
    For r = 1 To lastRow
        employeeId = CleanText_(CStrSafe_(data(r, DS_EMPLOYEE_ID)))
        fullName = NormalizeName_(CStrSafe_(data(r, DS_NAME)))

        If employeeId <> "" Then
            Dim rec As Object: Set rec = CreateObject("Scripting.Dictionary")
            rec.CompareMode = vbTextCompare

            rec("CitizenID") = CleanCCCD_(CStrSafe_(data(r, DS_CCCD)))
            rec("Department") = CleanText_(CStrSafe_(data(r, DS_DEPARTMENT)))
            rec("Section") = CleanText_(CStrSafe_(data(r, DS_SECTION)))
rec("Position") = CleanText_(CStrSafe_(data(r, DS_POSITION)))' Cập nhật mã nhân viêngDSByID(employeeId) = rec' Tối ưu logic: Nếu bị trùng tên, giữ lại người xuất hiện đầu tiên tránh đè nhầmIf fullName <> "" ThenIf Not gDSByName.Exists(fullName) ThenSet gDSByName(fullName) = recEnd IfEnd IfEnd IfNext rEnd SubPrivate Function GetDSRecord_(ByVal employeeId As String, ByVal fullName As String) As ObjectDim keyId As String: keyId = CleanText_(employeeId)If keyId <> "" And Not gDSByID Is Nothing ThenIf gDSByID.Exists(keyId) ThenSet GetDSRecord_ = gDSByID(keyId)Exit FunctionEnd IfEnd IfDim keyName As String: keyName = NormalizeName_(fullName)If keyName <> "" And Not gDSByName Is Nothing ThenIf gDSByName.Exists(keyName) ThenSet GetDSRecord_ = gDSByName(keyName)Exit FunctionEnd IfEnd IfSet GetDSRecord_ = NothingEnd FunctionPrivate Function BuildBatchPayload_( _ByRef data As Variant, _ByRef rowNums() As Long, _ByVal startIndex As Long, _ByVal endIndex As Long, _ByVal factory As String, _ByVal payMonth As String, _ByVal batchId As String, _ByVal batchNo As Long, _ByVal totalBatches As Long) As String' TỐI ƯU CỰC LỚN: Sử dụng Array String Buffer để chống phân mảnh RAM khi cộng chuỗiDim size As Long: size = (endIndex - startIndex + 1)Dim chunks() As String: ReDim chunks(1 To size)Dim idx As Long: idx = 1Dim i As LongFor i = startIndex To endIndexchunks(idx) = BuildEmployeeJson_(data, rowNums(i))idx = idx + 1Next iDim rowsJson As String: rowsJson = "[" & Join(chunks, ",") & "]"BuildBatchPayload_ = _"{""action"":""syncPayroll""," & _"""apiKey"":" & JsonString_(SYNC_API_KEY) & "," & _"""factory"":" & JsonString_(factory) & "," & _"""payMonth"":" & JsonString_(payMonth) & "," & _"""batchId"":" & JsonString_(batchId) & "," & _"""batchNo"":" & CStr(batchNo) & "," & _"""totalBatches"":" & CStr(totalBatches) & "," & _"""isFirstBatch"":" & IIf(batchNo = 1, "true", "false") & "," & _"""isLastBatch"":" & IIf(batchNo = totalBatches, "true", "false") & "," & _"""rows"":" & rowsJson & "}"End FunctionPrivate Function BuildEmployeeJson_(ByRef data As Variant, ByVal r As Long) As StringDim employeeId As String: employeeId = CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID)))Dim fullName As String: fullName = CleanText_(CStrSafe_(data(r, COL_NAME)))Dim ds As Object: Set ds = GetDSRecord_(employeeId, fullName)Dim citizenId As String, department As String, section As String, position As StringIf Not ds Is Nothing ThencitizenId = CleanCCCD_(CStrSafe_(ds("CitizenID")))department = CleanText_(CStrSafe_(ds("Department")))section = CleanText_(CStrSafe_(ds("Section")))position = CleanText_(CStrSafe_(ds("Position")))End IfDim j As String: j = "{"AddJsonString_ j, "EmployeeID", employeeId, TrueAddJsonString_ j, "FullName", fullName, TrueAddJsonString_ j, "CitizenID", citizenId, TrueAddJsonString_ j, "Department", department, TrueAddJsonString_ j, "Section", section, TrueAddJsonString_ j, "Position", position, TrueAddJsonNumber_ j, "TotalIncome", data(r, COL_TOTAL_INCOME), TrueAddJsonNumber_ j, "MonthlySalary", data(r, COL_MONTHLY_SALARY), TrueAddJsonNumber_ j, "BasicSalary", data(r, COL_BASIC), TrueAddJsonNumber_ j, "WorkingDays", data(r, COL_WORKING_DAYS), TrueAddJsonNumber_ j, "HolidayDays", data(r, COL_HOLIDAY_DAYS), TrueAddJsonNumber_ j, "PaidLeaveDays", data(r, COL_PAID_LEAVE), TrueAddJsonNumber_ j, "UnpaidLeaveDays", data(r, COL_UNPAID_LEAVE), TrueAddJsonNumber_ j, "RegionalMinimumLeaveDays", data(r, COL_MIN_WAGE_LEAVE), TrueAddJsonNumber_ j, "OvertimePay", data(r, COL_OT_PAY), TrueAddJsonNumber_ j, "OvertimeHours", data(r, COL_OT_HOURS), TrueAddJsonNumber_ j, "RestDayHours", data(r, COL_REST_HOURS), TrueAddJsonNumber_ j, "OvertimeRestDayHours", data(r, COL_OT_REST_HOURS), TrueAddJsonNumber_ j, "NightRestDayOvertimeHours", data(r, COL_NIGHT_REST_OT), TrueAddJsonNumber_ j, "HolidayHours", data(r, COL_HOLIDAY_HOURS), TrueAddJsonNumber_ j, "HolidayOvertimeHours", data(r, COL_HOLIDAY_OT), TrueAddJsonNumber_ j, "NightHolidayOvertimeHours", data(r, COL_NIGHT_HOLIDAY_OT), TrueAddJsonNumber_ j, "NightShiftDays", data(r, COL_NIGHT_DAYS), TrueAddJsonNumber_ j, "NightOvertimeHours", data(r, COL_NIGHT_OT), TrueAddJsonNumber_ j, "OtherMoney", data(r, COL_OTHER_MONEY), TrueAddJsonNumber_ j, "Discipline", data(r, COL_DISCIPLINE), TrueAddJsonNumber_ j, "Loyalty2Years", data(r, COL_LOYALTY2), TrueAddJsonNumber_ j, "Loyalty5Years", data(r, COL_LOYALTY5), TrueAddJsonNumber_ j, "Loyalty10Years", data(r, COL_LOYALTY10), TrueAddJsonNumber_ j, "Housing", data(r, COL_HOUSING), TrueAddJsonNumber_ j, "Transportation", data(r, COL_TRANSPORT), TrueAddJsonNumber_ j, "AttendanceBonus", data(r, COL_ATTENDANCE), TrueAddJsonNumber_ j, "SalesCommissionBonus", data(r, COL_COMMISSION), TrueAddJsonNumber_ j, "SeveranceUnusedLeave", data(r, COL_SEVERANCE), TrueAddJsonNumber_ j, "SocialInsurance", data(r, COL_SOCIAL), TrueAddJsonNumber_ j, "HealthInsurance", data(r, COL_HEALTH), TrueAddJsonNumber_ j, "UnemploymentInsurance", data(r, COL_UNEMPLOYMENT), TrueAddJsonNumber_ j, "PersonalIncomeTax", data(r, COL_TAX), TrueAddJsonNumber_ j, "Advance", data(r, COL_ADVANCE), TrueAddJsonNumber_ j, "OtherDeductions", data(r, COL_OTHER_DEDUCT), TrueAddJsonNumber_ j, "NetPay", data(r, COL_NET_PAY), FalseBuildEmployeeJson_ = j & "}"End FunctionPrivate Sub AddJsonString_(ByRef json As String, ByVal key As String, ByVal value As String, ByVal commaAfter As Boolean)json = json & JsonString_(key) & ":" & JsonString_(value)If commaAfter Then json = json & ","End SubPrivate Sub AddJsonNumber_(ByRef json As String, ByVal key As String, ByVal value As Variant, ByVal commaAfter As Boolean)json = json & JsonString_(key) & ":" & JsonNumber_(value)If commaAfter Then json = json & ","End SubPrivate Function JsonString_(ByVal s As String) As Strings = Replace(s, "", "\")s = Replace(s, """", """")s = Replace(s, vbCrLf, "\n")s = Replace(s, vbCr, "\n")s = Replace(s, vbLf, "\n")s = Replace(s, vbTab, "\t")JsonString_ = """" & s & """"End FunctionPrivate Function JsonNumber_(ByVal v As Variant) As StringIf IsError(v) Or IsEmpty(v) Or IsNull(v) Then JsonNumber_ = "0": Exit FunctionIf Not IsNumeric(v) Then JsonNumber_ = "0": Exit FunctionDim s As String: s = Format$(CDbl(v), "0.###############")s = Replace(s, Application.International(xlDecimalSeparator), ".")JsonNumber_ = sEnd Function' ============================================================' NETWORK CONNECTIVITY (REUSED OBJECT WITH RETRY)' ============================================================Private Function PostWithRetry_(ByVal payload As String, ByRef responseText As String) As BooleanDim attempt As Long, waitSeconds As LongFor attempt = 1 To MAX_RETRYresponseText = HttpPostJson_(payload)If ResponseOK_(responseText) ThenPostWithRetry_ = TrueExit FunctionEnd IfwaitSeconds = attempt * 2Application.StatusBar = "⚠️ Lỗi API mạng - Đang thử lại lần " & attempt & "/" & MAX_RETRY & " sau " & waitSeconds & " giây..."SleepSeconds_ waitSecondsNext attemptPostWithRetry_ = FalseEnd FunctionPrivate Function ResponseOK_(ByVal responseText As String) As BooleanResponseOK_ = InStr(1, responseText, """ok"":true", vbTextCompare) > 0End FunctionPrivate Function HttpPostJson_(ByVal payload As String) As StringOn Error GoTo EH' Nếu đối tượng HTTP bị hủy bất thường, tiến hành tạo lại ngầmIf m_Http Is Nothing Then Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")m_Http.Open "POST", GAS_URL, Falsem_Http.SetTimeouts 30000, 30000, 30000, 120000m_Http.SetRequestHeader "Content-Type", "text/plain;charset=utf-8"m_Http.Send payloadHttpPostJson_ = "HTTP " & m_Http.Status & ": " & m_Http.ResponseTextExit FunctionEH:HttpPostJson_ = "HTTP ERROR " & Err.Number & ": " & Err.DescriptionEnd Function' ============================================================' UTILITIES & STRING SANITIZATION' ============================================================Private Function CountSalaryRows_(ByRef data As Variant, ByVal lastRow As Long) As LongDim r As Long, count As LongFor r = 1 To lastRowIf CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID))) <> "" And _CleanText_(CStrSafe_(data(r, COL_NAME))) <> "" Thencount = count + 1End IfNext rCountSalaryRows_ = countEnd FunctionPrivate Function CStrSafe_(ByVal v As Variant) As StringIf IsError(v) Or IsNull(v) Or IsEmpty(v) Then CStrSafe_ = "" Else CStrSafe_ = CStr(v)End FunctionPrivate Function CleanText_(ByVal s As String) As Strings = Trim$(s)s = Replace(s, vbCr, " ")s = Replace(s, vbLf, " ")Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): LoopCleanText_ = sEnd FunctionPrivate Function CleanCCCD_(ByVal s As String) As Strings = CleanText_(s)Do While Left$(s, 1) = "'": s = Mid$(s, 2): LoopCleanCCCD_ = sEnd FunctionPrivate Function NormalizeName_(ByVal s As String) As StringNormalizeName_ = UCase$(CleanText_(s))End FunctionPrivate Function PreviousPayrollMonth_() As StringPreviousPayrollMonth_ = Format$(DateAdd("m", -1, Date), "mm-yyyy")End FunctionPrivate Function FileExists_(ByVal filePath As String) As BooleanFileExists_ = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)End FunctionPrivate Sub SleepSeconds_(ByVal seconds As Long)Dim endTime As Date: endTime = DateAdd("s", seconds, Now)Do While Now < endTimeDoEvents ' Giữ cho Excel không bị trạng thái Not Responding (Đơ)LoopEnd Sub
