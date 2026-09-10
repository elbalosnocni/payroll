Option Explicit

' ============================================================
' PAYROLL SYNC VBA - ĐÃ CHỈNH THEO FILE:
' PRINTING LINE 08-2026.xlsb
'
' FILE THỰC TẾ:
' Salary:
'   Họ tên = CF (84)
'   Mã NV  = D  (4)
'
' DSCNV:
'   Họ tên = B  (2)
'   CCCD  = H  (8)
'
' Theo file bạn gửi, các cột chức danh thực tế là:
'   AF (32) = BỘ PHẬN ENG / DEPARTMENT
'   AG (33) = CHỨC VỤ ENG / TITLE
'   AH (34) = DEPARTMENT / BỘ PHẬN VN
'   AI (35) = CHỨC VỤ / POSITION VN
'
' Vì yêu cầu hiển thị tiếng Việt, code lấy:
'   Department = AF
'   Section    = AH
'   Position   = AI
'
' Salary data bắt đầu thực tế từ row 9 trong file mẫu.
' Code không phụ thuộc tuyệt đối vào row 9:
' nó tìm dòng có Mã NV ở cột D và Họ tên ở CF.
'
' 2 XƯỞNG:
' Snack:
' \\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW\YYYY\SALARY MM-YYYY.xlsb
'
' Flexible:
' \\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line\YYYY\PRINTING LINE MM-YYYY.xlsb
'
' Password XLSB: 1234
'
' Ngày 10/09/2026 -> lấy tháng 08-2026.
'
' ============================================================

Private Const GAS_URL As String = _
"https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"

Private Const SYNC_API_KEY As String = "THAY_BANG_SYNC_API_KEY"

Private Const XLS_PASSWORD As String = "1234"

Private Const ROOT As String = "\\192.168.0.253\vn hr\"

' Salary columns
Private Const COL_EMPLOYEE_ID As Long = 4       'D
Private Const COL_BASIC As Long = 6             'F
Private Const COL_WORKING_DAYS As Long = 8       'H
Private Const COL_HOLIDAY_DAYS As Long = 9       'I
Private Const COL_PAID_LEAVE As Long = 10        'J
Private Const COL_UNPAID_LEAVE As Long = 11      'K
Private Const COL_OT_HOURS As Long = 12          'L
Private Const COL_REST_HOURS As Long = 13        'M
Private Const COL_OT_REST_HOURS As Long = 14     'N
Private Const COL_MIN_WAGE_LEAVE As Long = 15    'O
Private Const COL_NIGHT_REST_OT As Long = 16     'P
Private Const COL_NIGHT_OT As Long = 17          'Q
Private Const COL_HOLIDAY_HOURS As Long = 18     'R
Private Const COL_HOLIDAY_OT As Long = 19        'S
Private Const COL_NIGHT_HOLIDAY_OT As Long = 20  'T
Private Const COL_NIGHT_DAYS As Long = 21        'U

Private Const COL_OTHER_MONEY As Long = 24       'X
Private Const COL_DISCIPLINE As Long = 25        'Y
Private Const COL_LOYALTY2 As Long = 26          'Z
Private Const COL_LOYALTY5 As Long = 27          'AA
Private Const COL_LOYALTY10 As Long = 28         'AB
Private Const COL_HOUSING As Long = 29           'AC
Private Const COL_TRANSPORT As Long = 30         'AD
Private Const COL_ATTENDANCE As Long = 31        'AE
Private Const COL_SEVERANCE As Long = 32         'AF

Private Const COL_SOCIAL As Long = 33            'AG
Private Const COL_HEALTH As Long = 34            'AH
Private Const COL_UNEMPLOYMENT As Long = 35      'AI
Private Const COL_OTHER_DEDUCT As Long = 36      'AJ
Private Const COL_ADVANCE As Long = 37           'AK

Private Const COL_MONTHLY_SALARY As Long = 40    'AN
Private Const COL_OT_PAY As Long = 41            'AO
Private Const COL_COMMISSION As Long = 43        'AQ
Private Const COL_TOTAL_INCOME As Long = 44      'AR
Private Const COL_TAX As Long = 46               'AT
Private Const COL_NET_PAY As Long = 49           'AW
Private Const COL_NAME As Long = 84              'CF

' DSCNV
Private Const DS_NAME As Long = 2                 'B
Private Const DS_CCCD As Long = 8                 'H
Private Const DS_DEPARTMENT_EN As Long = 32       'AF
Private Const DS_TITLE_EN As Long = 33            'AG
Private Const DS_DEPARTMENT_VN As Long = 34       'AH
Private Const DS_POSITION_VN As Long = 35         'AI

' ============================================================
' ENTRY POINT
' ============================================================

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

    Debug.Print "Payroll month = " & payMonth
    Debug.Print snackPath
    Debug.Print flexPath

    If FileExists_(snackPath) Then
        SyncFactoryFile snackPath, "Snack", payMonth
    Else
        MsgBox "Không tìm thấy file Snack:" & vbCrLf & _
               snackPath, vbExclamation
    End If

    If FileExists_(flexPath) Then
        SyncFactoryFile flexPath, "Flexible", payMonth
    Else
        MsgBox "Không tìm thấy file Flexible:" & vbCrLf & _
               flexPath, vbExclamation
    End If

    MsgBox "Đã xử lý đồng bộ kỳ lương " & payMonth, vbInformation

End Sub

' Chạy riêng nếu muốn test 1 xưởng.
Public Sub SyncFlexibleOnly()

    Dim payMonth As String
    payMonth = PreviousPayrollMonth_()

    Dim y As String
    y = Right$(payMonth, 4)

    Dim p As String
    p = ROOT & _
        "SALARY - 2014 - 2015\Printing line\" & y & _
        "\PRINTING LINE " & payMonth & ".xlsb"

    If Not FileExists_(p) Then
        MsgBox "Không tìm thấy:" & vbCrLf & p, vbCritical
        Exit Sub
    End If

    SyncFactoryFile p, "Flexible", payMonth

End Sub

Public Sub SyncSnackOnly()

    Dim payMonth As String
    payMonth = PreviousPayrollMonth_()

    Dim y As String
    y = Right$(payMonth, 4)

    Dim p As String
    p = ROOT & _
        "SALARY - 2014 - 2015\VNLWW\" & y & _
        "\SALARY " & payMonth & ".xlsb"

    If Not FileExists_(p) Then
        MsgBox "Không tìm thấy:" & vbCrLf & p, vbCritical
        Exit Sub
    End If

    SyncFactoryFile p, "Snack", payMonth

End Sub

' ============================================================
' SYNC 1 WORKBOOK
' ============================================================

Public Sub SyncFactoryFile( _
    ByVal filePath As String, _
    ByVal factory As String, _
    ByVal payMonth As String)

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

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

    Dim dsMap As Object
    Set dsMap = BuildDSMap_(wsDS)

    Dim lastRow As Long
    lastRow = wsSalary.Cells(wsSalary.Rows.Count, COL_EMPLOYEE_ID).End(xlUp).Row

    Dim jsonRows As String
    jsonRows = "["

    Dim first As Boolean
    first = True

    Dim r As Long
    Dim countRows As Long

    For r = 1 To lastRow

        Dim employeeId As String
        Dim fullName As String

        employeeId = CleanText_(wsSalary.Cells(r, COL_EMPLOYEE_ID).Text)
        fullName = CleanText_(wsSalary.Cells(r, COL_NAME).Text)

        ' Chỉ lấy dòng nhân viên thật.
        ' Các dòng TOTAL/heading thường không có EmployeeID.
        If employeeId <> "" And fullName <> "" Then

            Dim ds As Object
            Set ds = Nothing

            If dsMap.Exists(NormalizeName_(fullName)) Then
                Set ds = dsMap(NormalizeName_(fullName))
            End If

            Dim citizenId As String
            Dim department As String
            Dim section As String
            Dim position As String

            citizenId = ""
            department = ""
            section = ""
            position = ""

            If Not ds Is Nothing Then
                citizenId = CleanCCCD_(CStr(ds("CitizenID")))
                department = CleanText_(CStr(ds("Department")))
                section = CleanText_(CStr(ds("Section")))
                position = CleanText_(CStr(ds("Position")))
            End If

            ' Nếu tên không match nhưng mã nhân viên có trong DSCNV,
            ' thử fallback theo EmployeeID.
            If citizenId = "" Then
                Dim dsByCode As Object
                Set dsByCode = FindDSByEmployeeCode_(wsDS, employeeId)

                If Not dsByCode Is Nothing Then
                    citizenId = CleanCCCD_(CStr(dsByCode("CitizenID")))
                    department = CleanText_(CStr(dsByCode("Department")))
                    section = CleanText_(CStr(dsByCode("Section")))
                    position = CleanText_(CStr(dsByCode("Position")))
                End If
            End If

            If Not first Then jsonRows = jsonRows & ","
            first = False

            jsonRows = jsonRows & "{"

            AddJsonString_ jsonRows, "EmployeeID", employeeId, True
            AddJsonString_ jsonRows, "FullName", fullName, True
            AddJsonString_ jsonRows, "CitizenID", citizenId, True
            AddJsonString_ jsonRows, "Department", department, True
            AddJsonString_ jsonRows, "Section", section, True
            AddJsonString_ jsonRows, "Position", position, True

            AddJsonNumber_ jsonRows, "TotalIncome", wsSalary.Cells(r, COL_TOTAL_INCOME).Value2, True
            AddJsonNumber_ jsonRows, "MonthlySalary", wsSalary.Cells(r, COL_MONTHLY_SALARY).Value2, True
            AddJsonNumber_ jsonRows, "BasicSalary", wsSalary.Cells(r, COL_BASIC).Value2, True
            AddJsonNumber_ jsonRows, "WorkingDays", wsSalary.Cells(r, COL_WORKING_DAYS).Value2, True
            AddJsonNumber_ jsonRows, "HolidayDays", wsSalary.Cells(r, COL_HOLIDAY_DAYS).Value2, True
            AddJsonNumber_ jsonRows, "PaidLeaveDays", wsSalary.Cells(r, COL_PAID_LEAVE).Value2, True
            AddJsonNumber_ jsonRows, "UnpaidLeaveDays", wsSalary.Cells(r, COL_UNPAID_LEAVE).Value2, True
            AddJsonNumber_ jsonRows, "RegionalMinimumLeaveDays", wsSalary.Cells(r, COL_MIN_WAGE_LEAVE).Value2, True

            AddJsonNumber_ jsonRows, "OvertimePay", wsSalary.Cells(r, COL_OT_PAY).Value2, True
            AddJsonNumber_ jsonRows, "OvertimeHours", wsSalary.Cells(r, COL_OT_HOURS).Value2, True
            AddJsonNumber_ jsonRows, "RestDayHours", wsSalary.Cells(r, COL_REST_HOURS).Value2, True
            AddJsonNumber_ jsonRows, "OvertimeRestDayHours", wsSalary.Cells(r, COL_OT_REST_HOURS).Value2, True
            AddJsonNumber_ jsonRows, "NightRestDayOvertimeHours", wsSalary.Cells(r, COL_NIGHT_REST_OT).Value2, True
            AddJsonNumber_ jsonRows, "HolidayHours", wsSalary.Cells(r, COL_HOLIDAY_HOURS).Value2, True
            AddJsonNumber_ jsonRows, "HolidayOvertimeHours", wsSalary.Cells(r, COL_HOLIDAY_OT).Value2, True
            AddJsonNumber_ jsonRows, "NightHolidayOvertimeHours", wsSalary.Cells(r, COL_NIGHT_HOLIDAY_OT).Value2, True
            AddJsonNumber_ jsonRows, "NightShiftDays", wsSalary.Cells(r, COL_NIGHT_DAYS).Value2, True
            AddJsonNumber_ jsonRows, "NightOvertimeHours", wsSalary.Cells(r, COL_NIGHT_OT).Value2, True

            AddJsonNumber_ jsonRows, "OtherMoney", wsSalary.Cells(r, COL_OTHER_MONEY).Value2, True
            AddJsonNumber_ jsonRows, "Discipline", wsSalary.Cells(r, COL_DISCIPLINE).Value2, True
            AddJsonNumber_ jsonRows, "Loyalty2Years", wsSalary.Cells(r, COL_LOYALTY2).Value2, True
            AddJsonNumber_ jsonRows, "Loyalty5Years", wsSalary.Cells(r, COL_LOYALTY5).Value2, True
            AddJsonNumber_ jsonRows, "Loyalty10Years", wsSalary.Cells(r, COL_LOYALTY10).Value2, True
            AddJsonNumber_ jsonRows, "Housing", wsSalary.Cells(r, COL_HOUSING).Value2, True
            AddJsonNumber_ jsonRows, "Transportation", wsSalary.Cells(r, COL_TRANSPORT).Value2, True
            AddJsonNumber_ jsonRows, "AttendanceBonus", wsSalary.Cells(r, COL_ATTENDANCE).Value2, True
            AddJsonNumber_ jsonRows, "SalesCommissionBonus", wsSalary.Cells(r, COL_COMMISSION).Value2, True
            AddJsonNumber_ jsonRows, "SeveranceUnusedLeave", wsSalary.Cells(r, COL_SEVERANCE).Value2, True

            AddJsonNumber_ jsonRows, "SocialInsurance", wsSalary.Cells(r, COL_SOCIAL).Value2, True
            AddJsonNumber_ jsonRows, "HealthInsurance", wsSalary.Cells(r, COL_HEALTH).Value2, True
            AddJsonNumber_ jsonRows, "UnemploymentInsurance", wsSalary.Cells(r, COL_UNEMPLOYMENT).Value2, True
            AddJsonNumber_ jsonRows, "PersonalIncomeTax", wsSalary.Cells(r, COL_TAX).Value2, True
            AddJsonNumber_ jsonRows, "Advance", wsSalary.Cells(r, COL_ADVANCE).Value2, True

            ' Khấu trừ khác = AJ
            AddJsonNumber_ jsonRows, "OtherDeductions", wsSalary.Cells(r, COL_OTHER_DEDUCT).Value2, True

            ' AW = NetPay
            AddJsonNumber_ jsonRows, "NetPay", wsSalary.Cells(r, COL_NET_PAY).Value2, False

            jsonRows = jsonRows & "}"

            countRows = countRows + 1

        End If

    Next r

    jsonRows = jsonRows & "]"

    Dim payload As String

    payload = "{"
    payload = payload & """action"":""syncPayroll"","
    payload = payload & """apiKey"":" & JsonString_(SYNC_API_KEY) & ","
    payload = payload & """factory"":" & JsonString_(factory) & ","
    payload = payload & """payMonth"":" & JsonString_(payMonth) & ","
    payload = payload & """rows"":" & jsonRows
    payload = payload & "}"

    Dim response As String
    response = HttpPostJson_(GAS_URL, payload)

    wb.Close SaveChanges:=False
    Set wb = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If InStr(1, response, """ok"":true", vbTextCompare) = 0 Then
        MsgBox _
            "Đồng bộ " & factory & " thất bại." & vbCrLf & _
            "Số dòng tạo: " & countRows & vbCrLf & _
            response, vbCritical
    Else
        MsgBox _
            "Đồng bộ " & factory & " thành công." & vbCrLf & _
            "Kỳ: " & payMonth & vbCrLf & _
            "Số nhân viên: " & countRows, vbInformation
    End If

    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False

    MsgBox _
        "Lỗi SyncFactoryFile:" & vbCrLf & _
        Err.Number & " - " & Err.Description, vbCritical

End Sub

' ============================================================
' DSCNV MAP
' ============================================================

Private Function BuildDSMap_(ws As Worksheet) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, DS_NAME).End(xlUp).Row

    Dim r As Long

    For r = 1 To lastRow

        Dim fullName As String
        fullName = CleanText_(ws.Cells(r, DS_NAME).Text)

        If fullName <> "" Then

            Dim employeeCode As String
            employeeCode = CleanText_(ws.Cells(r, 4).Text)

            If employeeCode <> "" Then

                Dim x As Object
                Set x = CreateObject("Scripting.Dictionary")

                x("EmployeeID") = employeeCode

                ' H là CCCD, phải lấy .Text để giữ số 0 đầu.
                x("CitizenID") = CleanCCCD_(ws.Cells(r, DS_CCCD).Text)

                ' Theo file thực tế đã upload:
                x("Department") = CleanText_(ws.Cells(r, DS_DEPARTMENT_EN).Text)

                ' Section = bộ phận VN
                x("Section") = CleanText_(ws.Cells(r, DS_DEPARTMENT_VN).Text)

                ' Position = chức vụ VN
                x("Position") = CleanText_(ws.Cells(r, DS_POSITION_VN).Text)

                d(NormalizeName_(fullName)) = x

            End If

        End If

    Next r

    Set BuildDSMap_ = d

End Function

Private Function FindDSByEmployeeCode_( _
    ws As Worksheet, _
    ByVal employeeId As String) As Object

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    Dim r As Long

    For r = 1 To lastRow

        If CleanText_(ws.Cells(r, 4).Text) = employeeId Then

            Dim x As Object
            Set x = CreateObject("Scripting.Dictionary")

            x("EmployeeID") = employeeId
            x("CitizenID") = CleanCCCD_(ws.Cells(r, DS_CCCD).Text)
            x("Department") = CleanText_(ws.Cells(r, DS_DEPARTMENT_EN).Text)
            x("Section") = CleanText_(ws.Cells(r, DS_DEPARTMENT_VN).Text)
            x("Position") = CleanText_(ws.Cells(r, DS_POSITION_VN).Text)

            Set FindDSByEmployeeCode_ = x
            Exit Function

        End If

    Next r

    Set FindDSByEmployeeCode_ = Nothing

End Function

' ============================================================
' JSON
' ============================================================

Private Sub AddJsonString_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As String, _
    ByVal commaAfter As Boolean)

    json = json & JsonString_(key) & ":" & JsonString_(value)

    If commaAfter Then json = json & ","

End Sub

Private Sub AddJsonNumber_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As Variant, _
    ByVal commaAfter As Boolean)

    json = json & JsonString_(key) & ":" & JsonNumber_(value)

    If commaAfter Then json = json & ","

End Sub

Private Function JsonString_(ByVal value As String) As String

    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbTab, "\t")

    JsonString_ = """" & value & """"

End Function

Private Function JsonNumber_(ByVal value As Variant) As String

    If IsError(value) Then
        JsonNumber_ = "0"
        Exit Function
    End If

    If IsEmpty(value) Then
        JsonNumber_ = "0"
        Exit Function
    End If

    If IsNumeric(value) Then

        Dim d As Double
        d = CDbl(value)

        If d = 0 Then
            JsonNumber_ = "0"
        Else
            ' VBA Decimal separator phụ thuộc Windows.
            JsonNumber_ = Replace( _
                Format$(d, "0.############"), _
                Application.International(xlDecimalSeparator), _
                ".")
        End If

    Else
        JsonNumber_ = "0"
    End If

End Function

' ============================================================
' HTTP
' ============================================================

Private Function HttpPostJson_( _
    ByVal url As String, _
    ByVal body As String) As String

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetRequestHeader _
        "Content-Type", _
        "text/plain;charset=utf-8"

    http.SetTimeouts 30000, 30000, 30000, 120000

    http.Send body

    If http.Status < 200 Or http.Status >= 300 Then

        HttpPostJson_ = _
            "HTTP " & http.Status & vbCrLf & _
            http.ResponseText

    Else

        HttpPostJson_ = http.ResponseText

    End If

End Function

' ============================================================
' DATE / FILE
' ============================================================

Private Function PreviousPayrollMonth_() As String

    Dim d As Date

    d = DateAdd("m", -1, Date)

    PreviousPayrollMonth_ = Format$(d, "mm-yyyy")

End Function

Private Function FileExists_(ByVal path As String) As Boolean

    On Error GoTo EH

    FileExists_ = (Len(Dir$(path)) > 0)

    Exit Function

EH:
    FileExists_ = False

End Function

' ============================================================
' TEXT
' ============================================================

Private Function CleanText_(ByVal value As String) As String

    value = Replace(value, ChrW(160), " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")

    CleanText_ = Trim$(value)

End Function

Private Function CleanCCCD_(ByVal value As String) As String

    value = Trim$(value)

    ' Excel có thể hiển thị dấu apostrophe khi nhập text.
    If Left$(value, 1) = "'" Then
        value = Mid$(value, 2)
    End If

    CleanCCCD_ = Trim$(value)

End Function

Private Function NormalizeName_(ByVal value As String) As String

    value = CleanText_(value)

    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop

    NormalizeName_ = UCase$(value)

End Function
