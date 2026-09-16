Option Explicit

Private Const API_URL As String = _
    "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"

Private Const SYNC_API_KEY As String = _
    "TRUONGCONGVU_SALARYSLIP_1988_Elbalosnocni"

Private Const FILE_PASSWORD As String = "1234"

Private Const ROOT_PATH As String = _
    "\\192.168.0.253\vn hr\SALARY - 2014 - 2015\"

Private Const MAX_RETRY As Long = 3

Private Const WAIT_SECONDS As Long = 5

Private Const SHEET_START_ROW As Long = 7

Private Const COL_EMPLOYEE_CODE As Long = 4
Private Const COL_BASIC_SALARY As Long = 6
Private Const COL_WORKING_DAYS As Long = 8
Private Const COL_HOLIDAY_DAYS As Long = 9
Private Const COL_PAID_LEAVE_DAYS As Long = 10
Private Const COL_UNPAID_LEAVE_DAYS As Long = 11
Private Const COL_OVERTIME_HOURS As Long = 12
Private Const COL_HOLIDAY_WORK_HOURS As Long = 13
Private Const COL_HOLIDAY_OVERTIME_HOURS As Long = 14
Private Const COL_MIN_REGIONAL_LEAVE_DAYS As Long = 15
Private Const COL_NIGHT_HOLIDAY_OT As Long = 16
Private Const COL_NIGHT_OT As Long = 17
Private Const COL_HOLIDAY_WORK_DAY_HOURS As Long = 18
Private Const COL_HOLIDAY_OT_DAY_HOURS As Long = 19
Private Const COL_NIGHT_SHIFT_DAYS As Long = 21
Private Const COL_OTHER_MONEY As Long = 24
Private Const COL_DISCIPLINARY_MONEY As Long = 25
Private Const COL_LOYALTY_2 As Long = 26
Private Const COL_LOYALTY_5 As Long = 27
Private Const COL_LOYALTY_10 As Long = 28
Private Const COL_HOUSING As Long = 29
Private Const COL_TRANSPORT As Long = 30
Private Const COL_ATTENDANCE As Long = 31
Private Const COL_SEVERANCE_UNUSED_LEAVE As Long = 32
Private Const COL_MONTHLY_SALARY As Long = 40
Private Const COL_OVERTIME_SALARY As Long = 41
Private Const COL_COMMISSION As Long = 43
Private Const COL_OTHER_INCOME As Long = 44
Private Const COL_OTHER_DEDUCTIONS As Long = 36
Private Const COL_ADVANCE As Long = 37
Private Const COL_SOCIAL_INSURANCE As Long = 33
Private Const COL_HEALTH_INSURANCE As Long = 34
Private Const COL_UNEMPLOYMENT_INSURANCE As Long = 35
Private Const COL_PERSONAL_INCOME_TAX As Long = 46
Private Const COL_NET_SALARY As Long = 49
Private Const COL_FULL_NAME As Long = 84

Private Const DSCNV_NAME_COL As Long = 2
Private Const DSCNV_CITIZEN_COL As Long = 8
Private Const DSCNV_DEPARTMENT_COL As Long = 32
Private Const DSCNV_SECTION_COL As Long = 33
Private Const DSCNV_POSITION_COL As Long = 34

Public Sub RunPayrollSync()

    Dim targetDate As Date
    Dim targetMonth As Date

    Dim salaryMonth As String

    Dim snackFile As String
    Dim flexibleFile As String

    Dim snackRecords As Collection
    Dim flexibleRecords As Collection
    Dim allRecords As Collection

    Dim payload As String
    Dim responseText As String

    Dim success As Boolean

    On Error GoTo ErrorHandler

    LogMessage "=== BAT DAU DONG BO LUONG ==="
    LogMessage "Thoi gian chay: " & _
               Format(Now, "dd/MM/yyyy HH:mm:ss")

    targetDate = Date
    targetMonth = DateAdd("m", -1, _
                          DateSerial(Year(targetDate), _
                                     Month(targetDate), _
                                     1))

    salaryMonth = _
        Format(targetMonth, "mm-yyyy")

    LogMessage "Thang luong can dong bo: " & _
               salaryMonth

    snackFile = FindSnackFile_(targetMonth)
    flexibleFile = FindFlexibleFile_(targetMonth)

    If Len(snackFile) = 0 Then
        LogMessage "Khong tim thay file Snack."
    Else
        LogMessage "Tim thay file Snack: " & snackFile
    End If

    If Len(flexibleFile) = 0 Then
        LogMessage "Khong tim thay file Flexible."
    Else
        LogMessage "Tim thay file Flexible: " & flexibleFile
    End If

    Set allRecords = New Collection

    If Len(snackFile) > 0 Then

        Set snackRecords = _
            ReadPayrollWorkbook_( _
                snackFile, _
                "Snack", _
                targetMonth)

        AppendCollection_ _
            allRecords, _
            snackRecords

    End If

    If Len(flexibleFile) > 0 Then

        Set flexibleRecords = _
            ReadPayrollWorkbook_( _
                flexibleFile, _
                "Flexible", _
                targetMonth)

        AppendCollection_ _
            allRecords, _
            flexibleRecords

    End If

    If allRecords.Count = 0 Then
        Err.Raise _
            vbObjectError + 1001, _
            "RunPayrollSync", _
            "Khong co du lieu luong de dong bo."
    End If

    LogMessage "Tong so dong du lieu: " & _
               CStr(allRecords.Count)

    payload = BuildSyncPayload_( _
        allRecords, _
        salaryMonth, _
        snackFile, _
        flexibleFile)

    LogMessage "Kich thuoc JSON: " & CStr(Len(payload)) & " ky tu."
    LogMessage "Dang gui du lieu len GAS..."

    success = PostWithRetry_( _
        payload, _
        responseText)

    If Not success Then
        Err.Raise _
            vbObjectError + 1002, _
            "RunPayrollSync", _
            "API dong bo that bai. Phan hoi: " & _
            responseText
    End If

    LogMessage "Dong bo thanh cong."
    LogMessage "Phan hoi API: " & responseText

    LogMessage "=== KET THUC DONG BO ==="

    Exit Sub

ErrorHandler:

    LogMessage "LOI: " & _
               Err.Number & _
               " - " & _
               Err.Description

    MsgBox _
        "Dong bo luong that bai." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Payroll Sync"

End Sub

Private Function FindSnackFile_( _
    ByVal targetMonth As Date) As String

    Dim yearFolder As String
    Dim fileName As String
    Dim fullPath As String

    yearFolder = _
        ROOT_PATH & _
        "VNLWW\" & _
        Format(targetMonth, "yyyy") & "\"

    fileName = _
        "SALARY " & _
        Format(targetMonth, "mm-yyyy") & _
        ".xlsb"

    fullPath = _
        yearFolder & fileName

    If FileExists_(fullPath) Then
        FindSnackFile_ = fullPath
    Else
        FindSnackFile_ = ""
    End If

End Function

Private Function FindFlexibleFile_( _
    ByVal targetMonth As Date) As String

    Dim yearFolder As String
    Dim fileName As String
    Dim fullPath As String

    yearFolder = _
        ROOT_PATH & _
        "Printing line\" & _
        Format(targetMonth, "yyyy") & "\"

    fileName = _
        "PRINTING LINE " & _
        Format(targetMonth, "mm-yyyy") & _
        ".xlsb"

    fullPath = _
        yearFolder & fileName

    If FileExists_(fullPath) Then
        FindFlexibleFile_ = fullPath
    Else
        FindFlexibleFile_ = ""
    End If

End Function

Private Function ReadPayrollWorkbook_( _
    ByVal filePath As String, _
    ByVal factoryName As String, _
    ByVal targetMonth As Date) As Collection

    Dim result As Collection
    Dim wb As Workbook
    Dim wsSalary As Worksheet
    Dim wsDSCNV As Worksheet

    Dim dscnvMap As Object

    Dim lastSalaryRow As Long
    Dim lastDSCNVRow As Long

    Dim rowIndex As Long

    Dim employeeName As String
    Dim normalizedName As String

    Dim employeeCode As String
    Dim citizenID As String
    Dim department As String
    Dim section As String
    Dim position As String

    Dim record As Object

    On Error GoTo ErrorHandler

    Set result = New Collection

    LogMessage _
        "[" & factoryName & _
        "] Dang mo file: " & _
        filePath

    Set wb = Workbooks.Open( _
        Filename:=filePath, _
        Password:=FILE_PASSWORD, _
        ReadOnly:=True, _
        UpdateLinks:=False, _
        AddToMru:=False)

    Set wsSalary = Nothing
    Set wsDSCNV = Nothing

    On Error Resume Next

    Set wsSalary = _
        wb.Worksheets("Salary")

    Set wsDSCNV = _
        wb.Worksheets("DSCNV")

    On Error GoTo ErrorHandler

    If wsSalary Is Nothing Then
        Err.Raise _
            vbObjectError + 2001, _
            "ReadPayrollWorkbook_", _
            "Khong tim thay sheet Salary."
    End If

    If wsDSCNV Is Nothing Then
        Err.Raise _
            vbObjectError + 2002, _
            "ReadPayrollWorkbook_", _
            "Khong tim thay sheet DSCNV."
    End If

    LogMessage _
        "[" & factoryName & _
        "] Dang tao index DSCNV..."

    Set dscnvMap = _
        BuildDSCNVMap_(wsDSCNV)

    lastSalaryRow = _
        LastUsedRow_( _
            wsSalary, _
            COL_FULL_NAME)

    lastDSCNVRow = _
        LastUsedRow_( _
            wsDSCNV, _
            DSCNV_NAME_COL)

    LogMessage _
        "[" & factoryName & _
        "] Salary last row: " & _
        CStr(lastSalaryRow)

    LogMessage _
        "[" & factoryName & _
        "] DSCNV last row: " & _
        CStr(lastDSCNVRow)

    For rowIndex = _
        SHEET_START_ROW To lastSalaryRow

        employeeName = _
            Trim$(CStr( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_FULL_NAME).Value2))

        If Len(employeeName) = 0 Then
            GoTo ContinueSalaryRow
        End If

        normalizedName = _
            NormalizeName_(employeeName)

        employeeCode = _
            Trim$(CStr( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_EMPLOYEE_CODE).Value2))

        If dscnvMap.Exists(normalizedName) Then

            citizenID = _
                dscnvMap(normalizedName)("CitizenID")

            department = _
                dscnvMap(normalizedName)("Department")

            section = _
                dscnvMap(normalizedName)("Section")

            position = _
                dscnvMap(normalizedName)("Position")

        Else

            citizenID = ""
            department = ""
            section = ""
            position = ""

            LogMessage _
                "[" & factoryName & _
                "] KHONG MAP DSCNV: " & _
                employeeName

        End If

        Set record = CreateObject("Scripting.Dictionary")

        record.Add _
            "employeeCode", _
            employeeCode

        record.Add _
            "fullName", _
            employeeName

        record.Add _
            "citizenID", _
            citizenID

        record.Add _
            "department", _
            department

        record.Add _
            "section", _
            section

        record.Add _
            "position", _
            position

        record.Add _
            "basicSalary", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_BASIC_SALARY))

        record.Add _
            "workingDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_WORKING_DAYS))

        record.Add _
            "holidayDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOLIDAY_DAYS))

        record.Add _
            "paidLeaveDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_PAID_LEAVE_DAYS))

        record.Add _
            "unpaidLeaveDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_UNPAID_LEAVE_DAYS))

        record.Add _
            "overtimeHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_OVERTIME_HOURS))

        record.Add _
            "holidayWorkHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOLIDAY_WORK_HOURS))

        record.Add _
            "holidayOvertimeHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOLIDAY_OVERTIME_HOURS))

        record.Add _
            "minimumRegionalLeaveDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_MIN_REGIONAL_LEAVE_DAYS))

        record.Add _
            "nightHolidayOvertimeHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_NIGHT_HOLIDAY_OT))

        record.Add _
            "nightOvertimeHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_NIGHT_OT))

        record.Add _
            "holidayWorkDayHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOLIDAY_WORK_DAY_HOURS))

        record.Add _
            "holidayOvertimeDayHours", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOLIDAY_OT_DAY_HOURS))

        record.Add _
            "nightShiftDays", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_NIGHT_SHIFT_DAYS))

        record.Add _
            "otherMoney", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_OTHER_MONEY))

        record.Add _
            "disciplinaryMoney", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_DISCIPLINARY_MONEY))

        record.Add _
            "loyalty2Years", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_LOYALTY_2))

        record.Add _
            "loyalty5Years", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_LOYALTY_5))

        record.Add _
            "loyalty10Years", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_LOYALTY_10))

        record.Add _
            "housingAllowance", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HOUSING))

        record.Add _
            "transportationAllowance", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_TRANSPORT))

        record.Add _
            "attendanceBonus", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_ATTENDANCE))

        record.Add _
            "severanceAndUnusedLeave", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_SEVERANCE_UNUSED_LEAVE))

        record.Add _
            "monthlySalary", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_MONTHLY_SALARY))

        record.Add _
            "overtimeSalary", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_OVERTIME_SALARY))

        record.Add _
            "commissionAndOverTargetBonus", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_COMMISSION))

        record.Add _
            "otherIncome", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_OTHER_INCOME))

        record.Add _
            "otherDeductions", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_OTHER_DEDUCTIONS))

        record.Add _
            "advancePayment", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_ADVANCE))

        record.Add _
            "socialInsurance", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_SOCIAL_INSURANCE))

        record.Add _
            "healthInsurance", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_HEALTH_INSURANCE))

        record.Add _
            "unemploymentInsurance", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_UNEMPLOYMENT_INSURANCE))

        record.Add _
            "personalIncomeTax", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_PERSONAL_INCOME_TAX))

        record.Add _
            "netSalary", _
            CellNumber_( _
                wsSalary.Cells( _
                    rowIndex, _
                    COL_NET_SALARY))

        result.Add record

ContinueSalaryRow:

    Next rowIndex

    wb.Close SaveChanges:=False
    Set wb = Nothing

    LogMessage _
        "[" & factoryName & _
        "] Doc duoc " & _
        CStr(result.Count) & _
        " dong."

    Set ReadPayrollWorkbook_ = result

    Exit Function

ErrorHandler:

    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If

    Err.Raise _
        Err.Number, _
        "ReadPayrollWorkbook_", _
        Err.Description

End Function

Private Function BuildDSCNVMap_( _
    ByVal ws As Worksheet) As Object

    Dim map As Object

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim employeeName As String
    Dim key As String

    Dim item As Object

    Set map = _
        CreateObject("Scripting.Dictionary")

    map.CompareMode = vbTextCompare

    lastRow = _
        LastUsedRow_( _
            ws, _
            DSCNV_NAME_COL)

    For rowIndex = _
        SHEET_START_ROW To lastRow

        employeeName = _
            Trim$(CStr( _
                ws.Cells( _
                    rowIndex, _
                    DSCNV_NAME_COL).Value2))

        If Len(employeeName) > 0 Then

            key = _
                NormalizeName_(employeeName)

            Set item = CreateObject("Scripting.Dictionary")

            item.Add _
                "CitizenID", _
                ReadCitizenID_( _
                    ws.Cells( _
                        rowIndex, _
                        DSCNV_CITIZEN_COL))

            item.Add _
                "Department", _
                Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        DSCNV_DEPARTMENT_COL).Value2))

            item.Add _
                "Section", _
                Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        DSCNV_SECTION_COL).Value2))

            item.Add _
                "Position", _
                Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        DSCNV_POSITION_COL).Value2))

            If map.Exists(key) Then

                LogMessage _
                    "CANH BAO: Trung ten trong DSCNV: " & _
                    employeeName

            Else

                map.Add key, item

            End If

        End If

    Next rowIndex

    Set BuildDSCNVMap_ = map

End Function

Private Function ReadCitizenID_( _
    ByVal cell As Range) As String

    Dim textValue As String
    Dim numberValue As Double
    Dim formatText As String

    On Error GoTo Fallback

    textValue = _
        Trim$(CStr(cell.Value2))

    textValue = _
        Replace(textValue, "'", "")

    textValue = _
        Replace(textValue, " ", "")

    If Len(textValue) = 0 Then
        ReadCitizenID_ = ""
        Exit Function
    End If

    If IsNumeric(cell.Value2) Then

        numberValue = _
            CDbl(cell.Value2)

        formatText = _
            CStr(cell.NumberFormat)

        If Len(formatText) > 0 _
           And InStr(1, formatText, "0", _
                    vbTextCompare) > 0 Then

            textValue = _
                Format$( _
                    numberValue, _
                    Replace( _
                        Replace( _
                            formatText, _
                            """", ""), _
                        " ", ""))

        Else

            If numberValue = _
               Fix(numberValue) Then

                textValue = _
                    Format$( _
                        numberValue, _
                        "0")

            Else

                textValue = _
                    CStr(cell.Text)

            End If

        End If

    End If

    If Right$(textValue, 2) = ".0" Then
        textValue = _
            Left$( _
                textValue, _
                Len(textValue) - 2)
    End If

    If IsNumeric(textValue) Then

        textValue = _
            Replace(textValue, ".", "")

        If Len(textValue) < 12 Then
            textValue = _
                Right$( _
                    String$(12, "0") & _
                    textValue, _
                    12)
        End If

    End If

    ReadCitizenID_ = textValue

    Exit Function

Fallback:

    ReadCitizenID_ = _
        Trim$(CStr(cell.Text))

End Function

Private Function NormalizeName_( _
    ByVal value As String) As String

    Dim textValue As String

    textValue = _
        Trim$(value)

    textValue = _
        Replace(textValue, vbTab, " ")

    Do While InStr(textValue, "  ") > 0

        textValue = _
            Replace( _
                textValue, _
                "  ", _
                " ")

    Loop

    NormalizeName_ = _
        LCase$(textValue)

End Function

Private Function CellNumber_( _
    ByVal cell As Range) As Double

    Dim value As Variant
    Dim textValue As String

    On Error GoTo SafeZero

    value = cell.Value2

    If IsNumeric(value) Then
        CellNumber_ = CDbl(value)
        Exit Function
    End If

    textValue = _
        Trim$(CStr(value))

    textValue = _
        Replace(textValue, ",", "")

    If IsNumeric(textValue) Then
        CellNumber_ = CDbl(textValue)
    Else
        CellNumber_ = 0
    End If

    Exit Function

SafeZero:

    CellNumber_ = 0

End Function

Private Function LastUsedRow_( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long) As Long

    LastUsedRow_ = _
        ws.Cells( _
            ws.Rows.Count, _
            columnNumber).End(xlUp).Row

End Function

Private Function FileExists_( _
    ByVal filePath As String) As Boolean

    FileExists_ = _
        (Len(Dir$(filePath, vbNormal)) > 0)

End Function

Private Sub AppendCollection_( _
    ByVal target As Collection, _
    ByVal source As Collection)

    Dim item As Variant

    If source Is Nothing Then
        Exit Sub
    End If

    For Each item In source
        target.Add item
    Next item

End Sub

Private Function BuildSyncPayload_( _
    ByVal records As Collection, _
    ByVal salaryMonth As String, _
    ByVal snackFile As String, _
    ByVal flexibleFile As String) As String

    Dim json As String
    Dim i As Long

    json = "{"

    json = json & _
        """action"":""syncPayroll"","

    json = json & _
        """apiKey"":""" & _
        JsonEscape_(SYNC_API_KEY) & _
        ""","

    json = json & _
        """salaryMonth"":""" & _
        JsonEscape_(salaryMonth) & _
        ""","

    json = json & _
        """source"":{"

    json = json & _
        """snackFile"":""" & _
        JsonEscape_(snackFile) & _
        ""","

    json = json & _
        """flexibleFile"":""" & _
        JsonEscape_(flexibleFile) & _
        """},"

    json = json & _
        """records"":["

    For i = 1 To records.Count

        If i > 1 Then
            json = json & ","
        End If

        json = json & _
            RecordToJson_(records(i))

    Next i

    json = json & "]"

    json = json & "}"

    BuildSyncPayload_ = json

End Function

Private Function RecordToJson_( _
    ByVal record As Object) As String

    Dim json As String

    json = "{"

    AddJsonString_ json, _
        "employeeCode", _
        record("employeeCode"), True

    AddJsonString_ json, _
        "fullName", _
        record("fullName"), True

    AddJsonString_ json, _
        "citizenID", _
        record("citizenID"), True

    AddJsonString_ json, _
        "department", _
        record("department"), True

    AddJsonString_ json, _
        "section", _
        record("section"), True

    AddJsonString_ json, _
        "position", _
        record("position"), True

    AddJsonNumber_ json, _
        "basicSalary", _
        record("basicSalary"), True

    AddJsonNumber_ json, _
        "workingDays", _
        record("workingDays"), True

    AddJsonNumber_ json, _
        "holidayDays", _
        record("holidayDays"), True

    AddJsonNumber_ json, _
        "paidLeaveDays", _
        record("paidLeaveDays"), True

    AddJsonNumber_ json, _
        "unpaidLeaveDays", _
        record("unpaidLeaveDays"), True

    AddJsonNumber_ json, _
        "overtimeHours", _
        record("overtimeHours"), True

    AddJsonNumber_ json, _
        "holidayWorkHours", _
        record("holidayWorkHours"), True

    AddJsonNumber_ json, _
        "holidayOvertimeHours", _
        record("holidayOvertimeHours"), True

    AddJsonNumber_ json, _
        "minimumRegionalLeaveDays", _
        record("minimumRegionalLeaveDays"), True

    AddJsonNumber_ json, _
        "nightHolidayOvertimeHours", _
        record("nightHolidayOvertimeHours"), True

    AddJsonNumber_ json, _
        "nightOvertimeHours", _
        record("nightOvertimeHours"), True

    AddJsonNumber_ json, _
        "holidayWorkDayHours", _
        record("holidayWorkDayHours"), True

    AddJsonNumber_ json, _
        "holidayOvertimeDayHours", _
        record("holidayOvertimeDayHours"), True

    AddJsonNumber_ json, _
        "nightShiftDays", _
        record("nightShiftDays"), True

    AddJsonNumber_ json, _
        "otherMoney", _
        record("otherMoney"), True

    AddJsonNumber_ json, _
        "disciplinaryMoney", _
        record("disciplinaryMoney"), True

    AddJsonNumber_ json, _
        "loyalty2Years", _
        record("loyalty2Years"), True

    AddJsonNumber_ json, _
        "loyalty5Years", _
        record("loyalty5Years"), True

    AddJsonNumber_ json, _
        "loyalty10Years", _
        record("loyalty10Years"), True

    AddJsonNumber_ json, _
        "housingAllowance", _
        record("housingAllowance"), True

    AddJsonNumber_ json, _
        "transportationAllowance", _
        record("transportationAllowance"), True

    AddJsonNumber_ json, _
        "attendanceBonus", _
        record("attendanceBonus"), True

    AddJsonNumber_ json, _
        "severanceAndUnusedLeave", _
        record("severanceAndUnusedLeave"), True

    AddJsonNumber_ json, _
        "monthlySalary", _
        record("monthlySalary"), True

    AddJsonNumber_ json, _
        "overtimeSalary", _
        record("overtimeSalary"), True

    AddJsonNumber_ json, _
        "commissionAndOverTargetBonus", _
        record("commissionAndOverTargetBonus"), True

    AddJsonNumber_ json, _
        "otherIncome", _
        record("otherIncome"), True

    AddJsonNumber_ json, _
        "otherDeductions", _
        record("otherDeductions"), True

    AddJsonNumber_ json, _
        "advancePayment", _
        record("advancePayment"), True

    AddJsonNumber_ json, _
        "socialInsurance", _
        record("socialInsurance"), True

    AddJsonNumber_ json, _
        "healthInsurance", _
        record("healthInsurance"), True

    AddJsonNumber_ json, _
        "unemploymentInsurance", _
        record("unemploymentInsurance"), True

    AddJsonNumber_ json, _
        "personalIncomeTax", _
        record("personalIncomeTax"), True

    AddJsonNumber_ json, _
        "netSalary", _
        record("netSalary"), False

    json = json & "}"

    RecordToJson_ = json

End Function

Private Sub AddJsonString_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As String, _
    ByVal addComma As Boolean)

    json = json & _
        """" & key & """:""" & _
        JsonEscape_(value) & """"

    If addComma Then
        json = json & ","
    End If

End Sub

Private Sub AddJsonNumber_( _
    ByRef json As String, _
    ByVal key As String, _
    ByVal value As Double, _
    ByVal addComma As Boolean)

    ' IMPORTANT: payroll numeric values are sent as JSON strings.
    ' This completely avoids locale-dependent VBA number serialization
    ' such as 0., .5, or 1,5 which are invalid JSON in some cases.
    ' Code.gs converts these strings back to real numbers before writing
    ' them into Google Sheets.
    json = json & _
        """ & key & "":"" & _
        JsonNumberString_(value) & """

    If addComma Then
        json = json & ","
    End If

End Sub

Private Function JsonNumberString_( _
    ByVal value As Double) As String

    Dim result As String

    If IsNumeric(value) Then
        result = Format$(value, "0.############################")
    Else
        result = "0"
    End If

    ' Force JSON-independent decimal point. Format$ above always supplies
    ' the leading zero because the integer part is mandatory in the format.
    If Application.International(xlDecimalSeparator) <> "." Then
        result = Replace$(result, _
            Application.International(xlDecimalSeparator), ".")
    End If

    If Len(result) = 0 Or result = "-" Or result = "." Or result = "-." Then
        result = "0"
    End If

    If Left$(result, 1) = "." Then
        result = "0" & result
    ElseIf Left$(result, 2) = "-." Then
        result = "-0" & Mid$(result, 2)
    End If

    JsonNumberString_ = result

End Function

Private Function JsonEscape_( _
    ByVal value As String) As String

    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String

    result = ""

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)

        If code < 0 Then code = code + 65536

        Select Case code
            Case 8
                result = result & "\b"
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 12
                result = result & "\f"
            Case 13
                result = result & "\r"
            Case 0 To 7, 11, 14 To 31
                result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case 34
                result = result & "\"""
            Case 92
                result = result & "\\"
            Case Else
                ' Encode non-ASCII UTF-16 code units as JSON \uXXXX.
                ' This keeps the entire request ASCII and avoids Excel/VBA
                ' code-page/UTF-8 conversion problems with Vietnamese text.
                If code > 127 Then
                    result = result & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i

    JsonEscape_ = result

End Function

Private Function PostWithRetry_( _
    ByVal payload As String, _
    ByRef responseText As String) As Boolean

    Dim attempt As Long
    Dim waitSeconds As Long

    For attempt = 1 To MAX_RETRY

        LogMessage _
            "API attempt " & _
            CStr(attempt) & _
            "/" & _
            CStr(MAX_RETRY)

        responseText = _
            HttpPostJson_(payload)

        LogMessage _
            "API response: " & _
            Left$(responseText, 1000)

        If ResponseOK_(responseText) Then

            PostWithRetry_ = True

            Exit Function

        End If

        If attempt < MAX_RETRY Then

            waitSeconds = _
                WAIT_SECONDS * attempt

            LogMessage _
                "API loi. Cho " & _
                CStr(waitSeconds) & _
                " giay roi thu lai."

            SleepSeconds_ waitSeconds

        End If

    Next attempt

    PostWithRetry_ = False

End Function

Private Function HttpPostJson_( _
    ByVal payload As String) As String

    Dim http As Object
    Dim bodyBytes As Variant

    On Error GoTo ErrorHandler

    Set http = _
        CreateObject( _
            "WinHttp.WinHttpRequest.5.1")

    bodyBytes = Utf8Bytes_(payload)

    http.Open _
        "POST", _
        API_URL, _
        False

    http.SetTimeouts _
        15000, _
        15000, _
        30000, _
        30000

    http.SetRequestHeader _
        "Content-Type", _
        "application/json; charset=utf-8"

    http.Send bodyBytes

    HttpPostJson_ = _
        CStr(http.ResponseText)

    Exit Function

ErrorHandler:

    HttpPostJson_ = _
        "{""success"":false,""error"":""" & _
        JsonEscape_(Err.Description) & _
        """}"

End Function

Private Function Utf8Bytes_( _
    ByVal textValue As String) As Variant

    Dim stream As Object
    Dim bytes As Variant
    Dim length As Long

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.Position = 0
    stream.Type = 1

    bytes = stream.Read
    length = UBound(bytes) - LBound(bytes) + 1

    ' ADODB.Stream writes UTF-8 BOM (EF BB BF).
    If length >= 3 Then
        If bytes(0) = &HEF And _
           bytes(1) = &HBB And _
           bytes(2) = &HBF Then

            stream.Position = 3
            bytes = stream.Read
        End If
    End If

    stream.Close
    Set stream = Nothing

    Utf8Bytes_ = bytes

End Function

Private Function ResponseOK_( _
    ByVal responseText As String) As Boolean

    Dim textValue As String

    textValue = _
        LCase$(Trim$(responseText))

    If InStr(1, textValue, _
             """success"":true", _
             vbTextCompare) > 0 Then

        ResponseOK_ = True

    Else

        ResponseOK_ = False

    End If

End Function

Private Sub SleepSeconds_( _
    ByVal seconds As Long)

    Dim endTime As Date

    endTime = _
        DateAdd( _
            "s", _
            seconds, _
            Now)

    Do While Now < endTime
        DoEvents
    Loop

End Sub

Private Sub LogMessage( _
    ByVal message As String)

    Dim logFile As String
    Dim fileNumber As Integer

    logFile = _
        ThisWorkbook.Path & _
        "\PayrollSync.log"

    fileNumber = _
        FreeFile

    Open logFile _
        For Append _
        As #fileNumber

    Print #fileNumber, _
        Format(Now, "yyyy-mm-dd HH:mm:ss") & _
        " | " & _
        message

    Close #fileNumber

End Sub
