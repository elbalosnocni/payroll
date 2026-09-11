Option Explicit

' ============================================================
' PAYROLL SYNC - VBA / Excel XLSB
' Optimized Batch Sync Engine
' ============================================================

Private Const GAS_URL As String = "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"
Private Const SYNC_API_KEY As String = "PUT_YOUR_RANDOM_SYNC_API_KEY_HERE"
Private Const XLS_PASSWORD As String = "1234"
Private Const ROOT As String = "\\192.168.0.253\vn hr\"

Private Const BATCH_SIZE As Long = 500
Private Const MAX_RETRY As Long = 3

' --- SALARY SHEET MAPPING ---
Private Const COL_EMPLOYEE_ID As Long = 4       ' D
Private Const COL_BASIC As Long = 6             ' F
Private Const COL_WORKING_DAYS As Long = 8      ' H
Private Const COL_HOLIDAY_DAYS As Long = 9      ' I
Private Const COL_PAID_LEAVE As Long = 10       ' J
Private Const COL_UNPAID_LEAVE As Long = 11     ' K
Private Const COL_OT_HOURS As Long = 12         ' L
Private Const COL_REST_HOURS As Long = 13       ' M
Private Const COL_OT_REST_HOURS As Long = 14    ' N
Private Const COL_MIN_WAGE_LEAVE As Long = 15   ' O
Private Const COL_NIGHT_REST_OT As Long = 16    ' P
Private Const COL_NIGHT_OT As Long = 17         ' Q
Private Const COL_HOLIDAY_HOURS As Long = 18    ' R
Private Const COL_HOLIDAY_OT As Long = 19       ' S
Private Const COL_NIGHT_HOLIDAY_OT As Long = 20 ' T
Private Const COL_NIGHT_DAYS As Long = 21       ' U
Private Const COL_OTHER_MONEY As Long = 24      ' X
Private Const COL_DISCIPLINE As Long = 25       ' Y
Private Const COL_LOYALTY2 As Long = 26         ' Z
Private Const COL_LOYALTY5 As Long = 27         ' AA
Private Const COL_LOYALTY10 As Long = 28        ' AB
Private Const COL_HOUSING As Long = 29          ' AC
Private Const COL_TRANSPORT As Long = 30        ' AD
Private Const COL_ATTENDANCE As Long = 31       ' AE
Private Const COL_SEVERANCE As Long = 32        ' AF
Private Const COL_SOCIAL As Long = 33           ' AG
Private Const COL_HEALTH As Long = 34           ' AH
Private Const COL_UNEMPLOYMENT As Long = 35     ' AI
Private Const COL_OTHER_DEDUCT As Long = 36     ' AJ
Private Const COL_ADVANCE As Long = 37          ' AK
Private Const COL_MONTHLY_SALARY As Long = 40   ' AN
Private Const COL_OT_PAY As Long = 41           ' AO
Private Const COL_COMMISSION As Long = 43       ' AQ
Private Const COL_TOTAL_INCOME As Long = 44     ' AR
Private Const COL_TAX As Long = 46              ' AT
Private Const COL_NET_PAY As Long = 49          ' AW
Private Const COL_NAME As Long = 84             ' CF

' --- DSCNV SHEET MAPPING ---
Private Const DS_NAME As Long = 2               ' B
Private Const DS_CCCD As Long = 8               ' H
Private Const DS_DEPARTMENT As Long = 32        ' AF
Private Const DS_SECTION As Long = 33           ' AG
Private Const DS_POSITION As Long = 34          ' AH

' Shared Network & Hash Dictionary Objects
Private gDSByName As Object
Private m_Http As Object

Public Sub SyncBothFactories()
    Dim payMonth As String: payMonth = PreviousPayrollMonth_()
    Dim y As String: y = Right$(payMonth, 4)

    Dim snackPath As String: snackPath = ROOT & "SALARY - 2014 - 2015\VNLWW\" & y & "\SALARY " & payMonth & ".xlsb"
    Dim flexPath As String: flexPath = ROOT & "SALARY - 2014 - 2015\Printing line\" & y & "\PRINTING LINE " & payMonth & ".xlsb"

    Dim msg As String: msg = "KỲ LƯƠNG ĐỒNG BỘ: " & payMonth & vbCrLf & vbCrLf
    Dim okSnack As Boolean, okFlex As Boolean

    Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")

    If FileExists_(snackPath) Then
        okSnack = SyncFactoryFile(snackPath, "Snack", payMonth)
        msg = msg & IIf(okSnack, "✓ Xưởng Snack: Thành công", "✗ Xưởng Snack: Thất bại") & vbCrLf
    Else
        msg = msg & "✗ Xưởng Snack: Không thấy file " & snackPath & vbCrLf
    End If

    If FileExists_(flexPath) Then
        okFlex = SyncFactoryFile(flexPath, "Flexible", payMonth)
        msg = msg & IIf(okFlex, "✓ Xưởng Flexible: Thành công", "✗ Xưởng Flexible: Thất bại") & vbCrLf
    Else
        msg = msg & "✗ Xưởng Flexible: Không thấy file " & flexPath & vbCrLf
    End If

    Set m_Http = Nothing
    Application.StatusBar = False
    MsgBox msg, IIf(okSnack Or okFlex, vbInformation, vbCritical), "Payroll Sync Final"
End Sub

Public Function SyncFactoryFile(ByVal filePath As String, ByVal factory As String, ByVal payMonth As String) As Boolean
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=filePath, Password:=XLS_PASSWORD, ReadOnly:=True, UpdateLinks:=0)

    Dim wsSalary As Worksheet: Set wsSalary = wb.Worksheets("Salary")
    Dim wsDS As Worksheet: Set wsDS = wb.Worksheets("DSCNV")

    Set gDSByName = CreateObject("Scripting.Dictionary")
    gDSByName.CompareMode = vbTextCompare
    BuildDSMaps_ wsDS

    Dim lastRow As Long
    lastRow = wsSalary.Cells(wsSalary.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row
    If lastRow < 1 Then GoTo FAIL_EXIT

    Dim salaryData As Variant
    salaryData = wsSalary.Range(wsSalary.Cells(1, 1), wsSalary.Cells(lastRow, COL_NAME)).Value2

    Dim totalRows As Long: totalRows = CountSalaryRows_(salaryData, lastRow)
    If totalRows = 0 Then GoTo FAIL_EXIT

    Dim rowNums() As Long: ReDim rowNums(1 To totalRows)
    Dim i As Long, r As Long, emp As String, nm As String
    i = 0
    For r = 1 To lastRow
        emp = CleanText_(CStrSafe_(salaryData(r, COL_EMPLOYEE_ID)))
        nm = CleanText_(CStrSafe_(salaryData(r, COL_NAME)))
        If emp <> "" And nm <> "" Then
            i = i + 1
            rowNums(i) = r
        End If
    Next r

    Dim totalBatches As Long: totalBatches = (totalRows + BATCH_SIZE - 1) \ BATCH_SIZE
    Dim batchId As String: batchId = factory & "_" & Replace(payMonth, "-", "") & "_" & Format$(Now, "yyyymmddhhnnss")
    Dim startIndex As Long, endIndex As Long, batchNo As Long: batchNo = 0

    For startIndex = 1 To totalRows Step BATCH_SIZE
        batchNo = batchNo + 1
        endIndex = startIndex + BATCH_SIZE - 1
        If endIndex > totalRows Then endIndex = totalRows

        Dim payload As String
        payload = BuildBatchPayload_(salaryData, rowNums, startIndex, endIndex, factory, payMonth, batchId, batchNo, totalBatches)

        Dim response As String
        If Not PostWithRetry_(payload, response) Then GoTo FAIL_EXIT
        If InStr(1, response, """ok"":true", vbTextCompare) = 0 Then GoTo FAIL_EXIT

        Application.StatusBar = factory & ": Đã gửi lô " & batchNo & "/" & totalBatches
    Next startIndex

    SyncFactoryFile = True
    GoTo CLEAN_EXIT

FAIL_EXIT:
    SyncFactoryFile = False
CLEAN_EXIT:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set gDSByName = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Function
EH:
    SyncFactoryFile = False
    Resume CLEAN_EXIT
End Function

Private Sub BuildDSMaps_(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, DS_NAME).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim data As Variant
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, DS_POSITION)).Value2

    Dim r As Long, fullName As String
    For r = 1 To lastRow
        fullName = NormalizeName_(CStrSafe_(data(r, DS_NAME)))
        If fullName <> "" Then
            If Not gDSByName.Exists(fullName) Then
                Dim rec As Object: Set rec = CreateObject("Scripting.Dictionary")
                rec("CitizenID") = CleanCCCD_(CStrSafe_(data(r, DS_CCCD)))
                rec("Department") = CleanText_(CStrSafe_(data(r, DS_DEPARTMENT)))
                rec("Section") = CleanText_(CStrSafe_(data(r, DS_SECTION)))
                rec("Position") = CleanText_(CStrSafe_(data(r, DS_POSITION)))
                Set gDSByName(fullName) = rec
            End If
        End If
    Next r
End Sub

Private Function BuildBatchPayload_(ByRef data As Variant, ByRef rowNums() As Long, ByVal startIndex As Long, ByVal endIndex As Long, ByVal factory As String, ByVal payMonth As String, ByVal batchId As String, ByVal batchNo As Long, ByVal totalBatches As Long) As String
    Dim size As Long: size = (endIndex - startIndex + 1)
    Dim chunks() As String: ReDim chunks(1 To size)
    Dim idx As Long: idx = 1
    Dim i As Long
    For i = startIndex To endIndex
        chunks(idx) = BuildEmployeeJson_(data, rowNums(i))
        idx = idx + 1
    Next i

    BuildBatchPayload_ = "{" & _
        """action"":""syncPayroll""," & _
        """apiKey"":" & JsonString_(SYNC_API_KEY) & "," & _
        """factory"":" & JsonString_(factory) & "," & _
        """payMonth"":" & JsonString_(payMonth) & "," & _
        """batchId"":" & JsonString_(batchId) & "," & _
        """batchNo"":" & CStr(batchNo) & "," & _
        """totalBatches"":" & CStr(totalBatches) & "," & _
        """isFirstBatch"":" & IIf(batchNo = 1, "true", "false") & "," & _
        """isLastBatch"":" & IIf(batchNo = totalBatches, "true", "false") & "," & _
        """rows"":[" & Join(chunks, ",") & "]}"
End Function

Private Function BuildEmployeeJson_(ByRef data As Variant, ByVal r As Long) As String
    Dim employeeId As String: employeeId = CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID)))
    Dim fullName As String: fullName = CleanText_(CStrSafe_(data(r, COL_NAME)))

    Dim citizenId As String, department As String, section As String, position As String
    Dim keyName As String: keyName = NormalizeName_(fullName)

    If gDSByName.Exists(keyName) Then
        Dim ds As Object: Set ds = gDSByName(keyName)
        citizenId = ds("CitizenID")
        department = ds("Department")
        section = ds("Section")
        position = ds("Position")
    End If

    Dim j As String: j = "{"
    j = j & """EmployeeID"":" & JsonString_(employeeId) & ","
    j = j & """FullName"":" & JsonString_(fullName) & ","
    j = j & """CitizenID"":" & JsonString_(citizenId) & ","
    j = j & """Department"":" & JsonString_(department) & ","
    j = j & """Section"":" & JsonString_(section) & ","
    j = j & """Position"":" & JsonString_(position) & ","
    j = j & """TotalIncome"":" & JsonNumber_(data(r, COL_TOTAL_INCOME)) & ","
    j = j & """MonthlySalary"":" & JsonNumber_(data(r, COL_MONTHLY_SALARY)) & ","
    j = j & """BasicSalary"":" & JsonNumber_(data(r, COL_BASIC)) & ","
    j = j & """WorkingDays"":" & JsonNumber_(data(r, COL_WORKING_DAYS)) & ","
    j = j & """HolidayDays"":" & JsonNumber_(data(r, COL_HOLIDAY_DAYS)) & ","
    j = j & """PaidLeaveDays"":" & JsonNumber_(data(r, COL_PAID_LEAVE)) & ","
    j = j & """UnpaidLeaveDays"":" & JsonNumber_(data(r, COL_UNPAID_LEAVE)) & ","
    j = j & """RegionalMinimumLeaveDays"":" & JsonNumber_(data(r, COL_MIN_WAGE_LEAVE)) & ","
    j = j & """OvertimePay"":" & JsonNumber_(data(r, COL_OT_PAY)) & ","
    j = j & """OvertimeHours"":" & JsonNumber_(data(r, COL_OT_HOURS)) & ","
    j = j & """RestDayHours"":" & JsonNumber_(data(r, COL_REST_HOURS)) & ","
    j = j & """OvertimeRestDayHours"":" & JsonNumber_(data(r, COL_OT_REST_HOURS)) & ","
    j = j & """NightRestDayOvertimeHours"":" & JsonNumber_(data(r, COL_NIGHT_REST_OT)) & ","
    j = j & """HolidayHours"":" & JsonNumber_(data(r, COL_HOLIDAY_HOURS)) & ","
    j = j & """HolidayOvertimeHours"":" & JsonNumber_(data(r, COL_HOLIDAY_OT)) & ","
    j = j & """NightHolidayOvertimeHours"":" & JsonNumber_(data(r, COL_NIGHT_HOLIDAY_OT)) & ","
    j = j & """NightShiftDays"":" & JsonNumber_(data(r, COL_NIGHT_DAYS)) & ","
    j = j & """NightOvertimeHours"":" & JsonNumber_(data(r, COL_NIGHT_OT)) & ","
    j = j & """OtherMoney"":" & JsonNumber_(data(r, COL_OTHER_MONEY)) & ","
    j = j & """Discipline"":" & JsonNumber_(data(r, COL_DISCIPLINE)) & ","
    j = j & """Loyalty2Years"":" & JsonNumber_(data(r, COL_LOYALTY2)) & ","
    j = j & """Loyalty5Years"":" & JsonNumber_(data(r, COL_LOYALTY5)) & ","
    j = j & """Loyalty10Years"":" & JsonNumber_(data(r, COL_LOYALTY10)) & ","
    j = j & """Housing"":" & JsonNumber_(data(r, COL_HOUSING)) & ","
    j = j & """Transportation"":" & JsonNumber_(data(r, COL_TRANSPORT)) & ","
    j = j & """AttendanceBonus"":" & JsonNumber_(data(r, COL_ATTENDANCE)) & ","
    j = j & """SalesCommissionBonus"":" & JsonNumber_(data(r, COL_COMMISSION)) & ","
    j = j & """SeveranceUnusedLeave"":" & JsonNumber_(data(r, COL_SEVERANCE)) & ","
    j = j & """SocialInsurance"":" & JsonNumber_(data(r, COL_SOCIAL)) & ","
    j = j & """HealthInsurance"":" & JsonNumber_(data(r, COL_HEALTH)) & ","
    j = j & """UnemploymentInsurance"":" & JsonNumber_(data(r, COL_UNEMPLOYMENT)) & ","
    j = j & """PersonalIncomeTax"":" & JsonNumber_(data(r, COL_TAX)) & ","
    j = j & """Advance"":" & JsonNumber_(data(r, COL_ADVANCE)) & ","
    j = j & """OtherDeductions"":" & JsonNumber_(data(r, COL_OTHER_DEDUCT)) & ","
    j = j & """NetPay"":" & JsonNumber_(data(r, COL_NET_PAY))
    BuildEmployeeJson_ = j & "}"
End Function

Private Function PostWithRetry_(ByVal payload As String, ByRef responseText As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To MAX_RETRY
        If m_Http Is Nothing Then Set m_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
        m_Http.Open "POST", GAS_URL, False
        m_Http.SetTimeouts 30000, 30000, 30000, 120000
        m_Http.SetRequestHeader "Content-Type", "text/plain;charset=utf-8"
        
        On Error Resume Next
        m_Http.Send payload
        If Err.Number = 0 Then
            responseText = m_Http.ResponseText
            If InStr(1, responseText, """ok"":true", vbTextCompare) > 0 Then
                PostWithRetry_ = True
                Exit Function
            End If
        End If
        On Error GoTo 0
        SleepSeconds_ attempt * 2
    Next attempt
    PostWithRetry_ = False
End Function

' --- Utility Methods ---
Private Function CountSalaryRows_(ByRef data As Variant, ByVal lastRow As Long) As Long
    Dim r As Long, cnt As Long
    For r = 1 To lastRow
        If CleanText_(CStrSafe_(data(r, COL_EMPLOYEE_ID))) <> "" And CleanText_(CStrSafe_(data(r, COL_NAME))) <> "" Then
            cnt = cnt + 1
        End If
    Next r
    CountSalaryRows_ = cnt
End Function

Private Function CStrSafe_(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then CStrSafe_ = "" Else CStrSafe_ = CStr(v)
End Function

Private Function CleanText_(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    CleanText_ = s
End Function

Private Function CleanCCCD_(ByVal s As String) As String
    s = CleanText_(s)
    Do While Left$(s, 1) = "'": s = Mid$(s, 2): Loop
    CleanCCCD_ = s
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

Private Function JsonString_(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonString_ = """" & s & """"
End Function

Private Function JsonNumber_(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Or Not IsNumeric(v) Then
        JsonNumber_ = "0"
    Else
        Dim s As String: s = Format$(CDbl(v), "0.###############")
        s = Replace(s, Application.International(xlDecimalSeparator), ".")
        JsonNumber_ = s
    End If
End Function

Private Sub SleepSeconds_(ByVal seconds As Long)
    Dim endTime As Date: endTime = DateAdd("s", seconds, Now)
    Do While Now < endTime: DoEvents: Loop
End Sub
