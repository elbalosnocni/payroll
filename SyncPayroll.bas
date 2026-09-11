Option Explicit

' ============================================================
' QUICK TEST
' Run this macro to verify the two expected source paths for
' the previous payroll month without opening the files.
' ============================================================
Public Sub CheckPreviousPayrollFiles()
    Dim pm As String
    pm = PreviousPayrollMonth_()

    Dim snackPath As String
    Dim flexPath As String

    snackPath = ROOT & "SALARY - 2014 - 2015\VNLWW\" & Right$(pm, 4) & _
                "\SALARY " & pm & ".xlsb"

    flexPath = ROOT & "SALARY - 2014 - 2015\Printing line\" & Right$(pm, 4) & _
               "\PRINTING LINE " & pm & ".xlsb"

    MsgBox "Kỳ lương: " & pm & vbCrLf & vbCrLf & _
           "Snack: " & IIf(FileExists_(snackPath), "OK", "KHÔNG TÌM THẤY") & vbCrLf & _
           snackPath & vbCrLf & vbCrLf & _
           "Flexible: " & IIf(FileExists_(flexPath), "OK", "KHÔNG TÌM THẤY") & vbCrLf & _
           flexPath, vbInformation, "Kiểm tra file Payroll"
End Sub


' ============================================================
' PAYROLL SYNC - VBA / Excel XLSB
' Version: Optimized batch sync
'
' SOURCE FILES
' Snack:
' \\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW\YYYY\SALARY MM-YYYY.xlsb
'
' Flexible:
' \\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line\YYYY\PRINTING LINE MM-YYYY.xlsb
'
' XLSB password: 1234
'
' DSCNV:
' B  = Họ tên
' H  = CCCD
' AF = Phòng ban
' AG = Bộ phận
' AH = Chức vụ
' D  = Mã NV (used as primary join key)
'
' Salary:
' D  = EmployeeID
' CF = FullName
' AR = TotalIncome
' AN = MonthlySalary
' F  = BasicSalary
' H/I/J/K/O = days
' AO/L/M/N/P/R/S/T/U/Q = OT
' X/Y/Z/AA/AB/AC/AD/AE/AQ/AF = income
' AG/AH/AI/AJ/AK/AT = deductions
' AW = NetPay
'
' IMPORTANT:
' 1) Replace SYNC_API_KEY with a NEW random key.
' 2) Put exactly the same key into GAS Script Properties.
' 3) The old key should be rotated because it was exposed in source/chat.
' ============================================================

Private Const GAS_URL As String = _
"https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"

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

' Module-level lookup maps
Private gDSByID As Object
Private gDSByName As Object

Public Sub SyncBothFactories()
    Dim payMonth As String
    payMonth = PreviousPayrollMonth_()

    Dim y As String
    y = Right$(payMonth, 4)

    Dim snackPath As String
    Dim flexPath As String

    snackPath = ROOT & _
        "SALARY - 2014 - 2015\VNLWW\" & y & _
        "\SALARY " & payMonth & ".xlsb"

    flexPath = ROOT & _
        "SALARY - 2014 - 2015\Printing line\" & y & _
        "\PRINTING LINE " & payMonth & ".xlsb"

    Dim msg As String
    Dim okSnack As Boolean, okFlex As Boolean

    msg = "KỲ LƯƠNG: " & payMonth & vbCrLf & vbCrLf

    If FileExists_(snackPath) Then
        okSnack = SyncFactoryFile(snackPath, "Snack", payMonth)
        If okSnack Then
            msg = msg & "✓ Snack: thành công" & vbCrLf
        Else
            msg = msg & "✗ Snack: thất bại" & vbCrLf
        End If
    Else
        msg = msg & "✗ Snack: không tìm thấy file" & vbCrLf
    End If

    If FileExists_(flexPath) Then
        okFlex = SyncFactoryFile(flexPath, "Flexible", payMonth)
        If okFlex Then
            msg = msg & "✓ Flexible: thành công" & vbCrLf
        Else
            msg = msg & "✗ Flexible: thất bại" & vbCrLf
        End If
    Else
        msg = msg & "✗ Flexible: không tìm thấy file" & vbCrLf
    End If

    Application.StatusBar = False
    MsgBox msg, IIf(okSnack Or okFlex, vbInformation, vbCritical), "Payroll Sync"
End Sub

Public Sub SyncSnackOnly()
    Dim pm As String
    pm = PreviousPayrollMonth_()

    Dim p As String
    p = ROOT & _
        "SALARY - 2014 - 2015\VNLWW\" & Right$(pm, 4) & _
        "\SALARY " & pm & ".xlsb"

    If Not FileExists_(p) Then
        MsgBox "Không tìm thấy file:" & vbCrLf & p, vbCritical
        Exit Sub
    End If

    Call SyncFactoryFile(p, "Snack", pm)
End Sub

Public Sub SyncFlexibleOnly()
    Dim pm As String
    pm = PreviousPayrollMonth_()

    Dim p As String
    p = ROOT & _
        "SALARY - 2014 - 2015\Printing line\" & Right$(pm, 4) & _
        "\PRINTING LINE " & pm & ".xlsb"

    If Not FileExists_(p) Then
        MsgBox "Không tìm thấy file:" & vbCrLf & p, vbCritical
        Exit Sub
    End If

    Call SyncFactoryFile(p, "Flexible", pm)
End Sub

Public Function SyncFactoryFile( _
    ByVal filePath As String, _
    ByVal factory As String, _
    ByVal payMonth As String) As Boolean

    On Error GoTo EH

    If SYNC_API_KEY = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE" Or Len(SYNC_API_KEY) < 24 Then
        MsgBox "Hãy thay SYNC_API_KEY bằng khóa ngẫu nhiên mới trước khi đồng bộ.", vbCritical
        Exit Function
    End If

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

    Dim wsSalary As Worksheet
    Dim wsDS As Worksheet

    Set wsSalary = wb.Worksheets("Salary")
    Set wsDS = wb.Worksheets("DSCNV")

    Set gDSByID = CreateObject("Scripting.Dictionary")
    Set gDSByName = CreateObject("Scripting.Dictionary")
    gDSByID.CompareMode = vbTextCompare
    gDSByName.CompareMode = vbTextCompare

    BuildDSMaps_ wsDS

    Dim lastRow As Long
    lastRow = wsSalary.Cells( _
        wsSalary.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox factory & ": không có dữ liệu Salary.", vbExclamation
        GoTo FAIL_EXIT
    End If

    ' Read Salary A:CF once for speed.
    Dim salaryData As Variant
    salaryData = wsSalary.Range( _
        wsSalary.Cells(1, 1), _
        wsSalary.Cells(lastRow, COL_NAME)).Value2

    Dim totalRows As Long
    totalRows = CountSalaryRows_(salaryData, lastRow)

    If totalRows = 0 Then
        MsgBox factory & ": không có nhân viên hợp lệ.", vbExclamation
        GoTo FAIL_EXIT
    End If

    Dim rowNums() As Long
    ReDim rowNums(1 To totalRows)

    Dim i As Long, r As Long
    Dim emp As String, nm As String

    i = 0

    For r = 2 To lastRow
        emp = CleanText_(CStrSafe_(salaryData(r, COL_EMPLOYEE_ID)))
        nm = CleanText_(CStrSafe_(salaryData(r, COL_NAME)))

        If emp <> "" And nm <> "" Then
            i = i + 1
            rowNums(i) = r
        End If
    Next r

    Dim totalBatches As Long
    totalBatches = (totalRows + BATCH_SIZE - 1) \ BATCH_SIZE

    ' One batch ID for the entire factory/month sync.
    Dim batchId As String
    batchId = factory & "_" & _
              Replace(payMonth, "-", "") & "_" & _
              Format$(Now, "yyyymmddhhnnss")

    Dim startIndex As Long
    Dim endIndex As Long
    Dim batchNo As Long

    For startIndex = 1 To totalRows Step BATCH_SIZE
        batchNo = batchNo + 1

        endIndex = startIndex + BATCH_SIZE - 1
        If endIndex > totalRows Then endIndex = totalRows

        Dim payload As String
        payload = BuildBatchPayload_( _
            salaryData, rowNums, _
            startIndex, endIndex, _
            factory, payMonth, _
            batchId, batchNo, totalBatches)

        Dim response As String

        If Not PostWithRetry_(payload, response) Then
            MsgBox factory & _
                " thất bại tại batch " & batchNo & "/" & totalBatches & _
                vbCrLf & response, vbCritical
            GoTo FAIL_EXIT
        End If

        If Not ResponseOK_(response) Then
            MsgBox factory & _
                " server từ chối batch " & batchNo & "/" & totalBatches & _
                vbCrLf & response, vbCritical
            GoTo FAIL_EXIT
        End If

        Application.StatusBar = _
            factory & " - batch " & batchNo & "/" & totalBatches & _
            " - " & endIndex & "/" & totalRows & " nhân viên"
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

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Function

EH:
    SyncFactoryFile = False

    MsgBox _
        "Lỗi SyncFactoryFile:" & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    Resume CLEAN_EXIT
End Function

Private Sub BuildDSMaps_(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells( _
        ws.Rows.Count, DS_EMPLOYEE_ID).End(xlUp).Row

    If lastRow < 1 Then Exit Sub

    ' Read A:AH once.
    Dim data As Variant
    data = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastRow, DS_POSITION)).Value2

    Dim r As Long
    Dim employeeId As String
    Dim fullName As String

    For r = 2 To lastRow
        employeeId = CleanText_(CStrSafe_(data(r, DS_EMPLOYEE_ID)))
        fullName = NormalizeName_(CStrSafe_(data(r, DS_NAME)))

        If employeeId <> "" Then
            Dim rec As Object
            Set rec = CreateObject("Scripting.Dictionary")
            rec.CompareMode = vbTextCompare

            rec("CitizenID") = CleanCCCD_(CStrSafe_(data(r, DS_CCCD)))
            rec("Department") = CleanText_(CStrSafe_(data(r, DS_DEPARTMENT)))
            rec("Section") = CleanText_(CStrSafe_(data(r, DS_SECTION)))
            rec("Position") = CleanText_(CStrSafe_(data(r, DS_POSITION)))

            ' Employee ID is the primary key.
            If gDSByID.Exists(employeeId) Then
                gDSByID.Remove employeeId
            End If
            gDSByID.Add employeeId, rec

            ' Name is fallback only.
            If fullName <> "" Then
                If Not gDSByName.Exists(fullName) Then
                    gDSByName.Add fullName, rec
                End If
            End If
        End If
    Next r
End Sub

Private Function GetDSRecord_( _
    ByVal employeeId As String, _
    ByVal fullName As String) As Object

    Dim keyId As String
    keyId = CleanText_(employeeId)

    If keyId <> "" Then
        If Not gDSByID Is Nothing Then
            If gDSByID.Exists(keyId) Then
                Set GetDSRecord_ = gDSByID(keyId)
                Exit Function
            End If
        End If
    End If

    Dim keyName As String
    keyName = NormalizeName_(fullName)

    If keyName <> "" Then
        If Not gDSByName Is Nothing Then
            If gDSByName.Exists(keyName) Then
                Set GetDSRecord_ = gDSByName(keyName)
                Exit Function
            End If
        End If
    End If

    Set GetDSRecord_ = Nothing
End Function

Private Function BuildBatchPayload_( _
    ByRef data As Variant, _
    ByRef rowNums() As Long, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long, _
    ByVal factory As String, _
    ByVal payMonth As String, _
    ByVal batchId As String, _
    ByVal batchNo As Long, _
    ByVal totalBatches As Long) As String

    Dim rowsJson As String
    rowsJson = "["

    Dim first As Boolean
    first = True

    Dim i As Long, r As Long

    For i = startIndex To endIndex
        r = rowNums(i)

        If Not first Then rowsJson = rowsJson & ","
        first = False

        rowsJson = rowsJson & BuildEmployeeJson_(data, r)
    Next i

    rowsJson = rowsJson & "]"

    BuildBatchPayload_ = _
        "{""action"":""syncPayroll""," & _
        """apiKey"":" & JsonString_(SYNC_API_KEY) & "," & _
        """factory"":" & JsonString_(factory) & "," & _
        """payMonth"":" & JsonString_(payMonth) & "," & _
        """batchId"":" & JsonString_(batchId) & "," & _
        """batchNo"":" & CStr(batchNo) & "," & _
        """totalBatches"":" & CStr(totalBatches) & "," & _
        """isFirstBatch"":" & IIf(batchNo = 1, "true", "false") & "," & _
        """isLastBatch"":" & IIf(batchNo = totalBatches, "true", "false") & "," & _
        """rows"":" & rowsJson & "}"
End Function

Private Function BuildEmployeeJson_( _
    ByRef data As Variant, _
    ByVal r As Long) As String

    Dim employeeId As String
    Dim fullName As String

    employeeId = CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID)))
    fullName = CleanText_(CStrSafe_(data(r, COL_NAME)))

    Dim ds As Object
    Set ds = GetDSRecord_(employeeId, fullName)

    Dim citizenId As String
    Dim department As String
    Dim section As String
    Dim position As String

    If Not ds Is Nothing Then
        citizenId = CleanCCCD_(CStrSafe_(ds("CitizenID")))
        department = CleanText_(CStrSafe_(ds("Department")))
        section = CleanText_(CStrSafe_(ds("Section")))
        position = CleanText_(CStrSafe_(ds("Position")))
    End If

    Dim j As String
    j = "{"

    AddJsonString_ j, "EmployeeID", employeeId, True
    AddJsonString_ j, "FullName", fullName, True
    AddJsonString_ j, "CitizenID", citizenId, True
    AddJsonString_ j, "Department", department, True
    AddJsonString_ j, "Section", section, True
    AddJsonString_ j, "Position", position, True

    AddJsonNumber_ j, "TotalIncome", data(r, COL_TOTAL_INCOME), True
    AddJsonNumber_ j, "MonthlySalary", data(r, COL_MONTHLY_SALARY), True
    AddJsonNumber_ j, "BasicSalary", data(r, COL_BASIC), True
    AddJsonNumber_ j, "WorkingDays", data(r, COL_WORKING_DAYS), True
    AddJsonNumber_ j, "HolidayDays", data(r, COL_HOLIDAY_DAYS), True
    AddJsonNumber_ j, "PaidLeaveDays", data(r, COL_PAID_LEAVE), True
    AddJsonNumber_ j, "UnpaidLeaveDays", data(r, COL_UNPAID_LEAVE), True
    AddJsonNumber_ j, "RegionalMinimumLeaveDays", data(r, COL_MIN_WAGE_LEAVE), True

    AddJsonNumber_ j, "OvertimePay", data(r, COL_OT_PAY), True
    AddJsonNumber_ j, "OvertimeHours", data(r, COL_OT_HOURS), True
    AddJsonNumber_ j, "RestDayHours", data(r, COL_REST_HOURS), True
    AddJsonNumber_ j, "OvertimeRestDayHours", data(r, COL_OT_REST_HOURS), True
    AddJsonNumber_ j, "NightRestDayOvertimeHours", data(r, COL_NIGHT_REST_OT), True
    AddJsonNumber_ j, "HolidayHours", data(r, COL_HOLIDAY_HOURS), True
    AddJsonNumber_ j, "HolidayOvertimeHours", data(r, COL_HOLIDAY_OT), True
    AddJsonNumber_ j, "NightHolidayOvertimeHours", data(r, COL_NIGHT_HOLIDAY_OT), True
    AddJsonNumber_ j, "NightShiftDays", data(r, COL_NIGHT_DAYS), True
    AddJsonNumber_ j, "NightOvertimeHours", data(r, COL_NIGHT_OT), True

    AddJsonNumber_ j, "OtherMoney", data(r, COL_OTHER_MONEY), True
    AddJsonNumber_ j, "Discipline", data(r, COL_DISCIPLINE), True
    AddJsonNumber_ j, "Loyalty2Years", data(r, COL_LOYALTY2), True
    AddJsonNumber_ j, "Loyalty5Years", data(r, COL_LOYALTY5), True
    AddJsonNumber_ j, "Loyalty10Years", data(r, COL_LOYALTY10), True
    AddJsonNumber_ j, "Housing", data(r, COL_HOUSING), True
    AddJsonNumber_ j, "Transportation", data(r, COL_TRANSPORT), True
    AddJsonNumber_ j, "AttendanceBonus", data(r, COL_ATTENDANCE), True
    AddJsonNumber_ j, "SalesCommissionBonus", data(r, COL_COMMISSION), True
    AddJsonNumber_ j, "SeveranceUnusedLeave", data(r, COL_SEVERANCE), True

    AddJsonNumber_ j, "SocialInsurance", data(r, COL_SOCIAL), True
    AddJsonNumber_ j, "HealthInsurance", data(r, COL_HEALTH), True
    AddJsonNumber_ j, "UnemploymentInsurance", data(r, COL_UNEMPLOYMENT), True
    AddJsonNumber_ j, "PersonalIncomeTax", data(r, COL_TAX), True
    AddJsonNumber_ j, "Advance", data(r, COL_ADVANCE), True
    AddJsonNumber_ j, "OtherDeductions", data(r, COL_OTHER_DEDUCT), True

    AddJsonNumber_ j, "NetPay", data(r, COL_NET_PAY), False

    j = j & "}"
    BuildEmployeeJson_ = j
End Function

Private Sub AddJsonString_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As String, _
    ByVal commaAfter As Boolean)

    json = json & _
        JsonString_(key) & ":" & JsonString_(value)

    If commaAfter Then json = json & ","
End Sub

Private Sub AddJsonNumber_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As Variant, _
    ByVal commaAfter As Boolean)

    json = json & _
        JsonString_(key) & ":" & JsonNumber_(value)

    If commaAfter Then json = json & ","
End Sub

Private Function JsonString_(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\"""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, ChrW(8), "\b")
    s = Replace(s, ChrW(12), "\f")

    JsonString_ = """" & s & """"
End Function

Private Function JsonNumber_(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        JsonNumber_ = "0"
        Exit Function
    End If

    If Not IsNumeric(v) Then
        JsonNumber_ = "0"
        Exit Function
    End If

    Dim s As String
    s = Format$(CDbl(v), "0.###############")

    ' JSON requires dot as decimal separator.
    s = Replace( _
        s, _
        Application.International(xlDecimalSeparator), _
        ".")

    JsonNumber_ = s
End Function

Private Function PostWithRetry_( _
    ByVal payload As String, _
    ByRef responseText As String) As Boolean

    Dim attempt As Long
    Dim waitSeconds As Long

    For attempt = 1 To MAX_RETRY
        responseText = HttpPostJson_(payload)

        If ResponseOK_(responseText) Then
            PostWithRetry_ = True
            Exit Function
        End If

        waitSeconds = attempt * 2
        Application.StatusBar = _
            "API lỗi - thử lại " & attempt & "/" & MAX_RETRY & _
            " sau " & waitSeconds & " giây..."

        SleepSeconds_ waitSeconds
    Next attempt

    PostWithRetry_ = False
End Function

Private Function ResponseOK_(ByVal responseText As String) As Boolean
    ResponseOK_ = _
        InStr(1, responseText, """ok"":true", vbTextCompare) > 0
End Function

Private Function HttpPostJson_(ByVal payload As String) As String
    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", GAS_URL, False

    http.SetTimeouts _
        30000, _
        30000, _
        30000, _
        120000

    http.SetRequestHeader _
        "Content-Type", _
        "text/plain;charset=utf-8"

    http.Send payload

    HttpPostJson_ = _
        "HTTP " & http.Status & ": " & http.ResponseText

    Exit Function

EH:
    HttpPostJson_ = _
        "HTTP ERROR " & Err.Number & ": " & Err.Description
End Function

Private Function CountSalaryRows_( _
    ByRef data As Variant, _
    ByVal lastRow As Long) As Long

    Dim r As Long
    Dim count As Long

    If lastRow < 2 Then
        CountSalaryRows_ = 0
        Exit Function
    End If

    For r = 2 To lastRow
        If CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID))) <> "" And _
           CleanText_(CStrSafe_(data(r, COL_NAME))) <> "" Then
            count = count + 1
        End If
    Next r

    CountSalaryRows_ = count
End Function

Private Function CStrSafe_(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CStrSafe_ = ""
    Else
        CStrSafe_ = CStr(v)
    End If
End Function

Private Function CleanText_(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanText_ = s
End Function

Private Function CleanCCCD_(ByVal s As String) As String
    s = CleanText_(s)

    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop

    ' CCCD is 12 digits. If the source cell was accidentally converted
    ' to a numeric 11-digit value, restore the leading zero.
    If Len(s) = 11 And IsDigitsOnly_(s) Then
        s = "0" & s
    End If

    CleanCCCD_ = s
End Function

Private Function IsDigitsOnly_(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i

    IsDigitsOnly_ = True
End Function

Private Function NormalizeName_(ByVal s As String) As String
    NormalizeName_ = UCase$(CleanText_(s))
End Function

Private Function PreviousPayrollMonth_() As String
    PreviousPayrollMonth_ = Format$(DateAdd("m", -1, Date), "mm-yyyy")
End Function

Private Function FileExists_(ByVal filePath As String) As Boolean
    FileExists_ = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
End Function

Private Sub SleepSeconds_(ByVal seconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", seconds, Now)

    Do While Now < endTime
        DoEvents
    Loop
End Sub
