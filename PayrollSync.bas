Option Explicit

Private Const PAYROLL_SYNC_VERSION As String = "2.3-HEADER-MAPPING-ROBUST"

Private Const API_URL As String = _
    "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"

Private Const SYNC_API_KEY As String = _
    "TRUONGCONGVU_SALARYSLIP_1988_Elbalosnocni"

Private Const FILE_PASSWORD As String = "2410"

Private Const ROOT_PATH As String = _
    "\\192.168.0.253\vn hr\SALARY - 2014 - 2015\"

Private Const MAX_RETRY As Long = 3

Private Const WAIT_SECONDS As Long = 5

Private Const SHEET_START_ROW As Long = 7

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

    If allRecords.count = 0 Then
        Err.Raise _
            vbObjectError + 1001, _
            "RunPayrollSync", _
            "Khong co du lieu luong de dong bo."
    End If

    LogMessage "Tong so dong du lieu: " & _
               CStr(allRecords.count)

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
    Dim salaryMap As Object
    Dim extraHeaders As Object

    Dim lastSalaryRow As Long
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

    LogMessage "[" & factoryName & "] Dang mo file: " & filePath

    Set wb = Workbooks.Open( _
        fileName:=filePath, _
        password:=FILE_PASSWORD, _
        ReadOnly:=True, _
        UpdateLinks:=False, _
        AddToMru:=False)

    On Error Resume Next
    Set wsSalary = wb.Worksheets("Salary")
    Set wsDSCNV = wb.Worksheets("DSCNV")
    On Error GoTo ErrorHandler

    If wsSalary Is Nothing Then
        Err.Raise vbObjectError + 2001, "ReadPayrollWorkbook_", _
                  "Khong tim thay sheet Salary."
    End If

    If wsDSCNV Is Nothing Then
        Err.Raise vbObjectError + 2002, "ReadPayrollWorkbook_", _
                  "Khong tim thay sheet DSCNV."
    End If

    LogMessage "[" & factoryName & "] Dang nhan dien Header sheet Salary..."
    Set salaryMap = BuildSalaryHeaderMap_(wsSalary)

    LogMessage "[" & factoryName & "] Dang nhan dien Header sheet DSCNV..."
    Set dscnvMap = BuildDSCNVMap_(wsDSCNV)

    lastSalaryRow = FindPayrollDataLastRow_(wsSalary, salaryMap("fullName"))

    LogMessage "[" & factoryName & "] Salary data last row: " & CStr(lastSalaryRow)

    Set extraHeaders = CollectExtraHeaders_(wsSalary, salaryMap)

    For rowIndex = SHEET_START_ROW To lastSalaryRow

        employeeName = CellText_(wsSalary.Cells(rowIndex, salaryMap("fullName")))
        employeeCode = CellText_(wsSalary.Cells(rowIndex, salaryMap("employeeCode")))

        If Len(employeeName) = 0 Or Len(employeeCode) = 0 Then GoTo ContinueSalaryRow
        If IsSummaryRow_(employeeName, employeeCode) Then GoTo ContinueSalaryRow

        normalizedName = NormalizeName_(employeeName)

        Dim dscnvItem As Object
        Set dscnvItem = Nothing

        If Len(employeeCode) > 0 Then
            If dscnvMap.Exists("CODE|" & UCase$(employeeCode)) Then
                Set dscnvItem = dscnvMap("CODE|" & UCase$(employeeCode))
            End If
        End If

        If dscnvItem Is Nothing Then
            If dscnvMap.Exists("NAME|" & normalizedName) Then
                Set dscnvItem = dscnvMap("NAME|" & normalizedName)
            End If
        End If

        If Not dscnvItem Is Nothing Then
            citizenID = dscnvItem("CitizenID")
            department = dscnvItem("Department")
            section = dscnvItem("Section")
            position = dscnvItem("Position")
        Else
            citizenID = ""
            department = ""
            section = ""
            position = ""
            LogMessage "[" & factoryName & "] KHONG MAP DSCNV: " & employeeName & " / " & employeeCode
        End If

        Set record = CreateObject("Scripting.Dictionary")
        record.CompareMode = vbTextCompare

        record.Add "employeeCode", employeeCode
        record.Add "fullName", employeeName
        record.Add "citizenID", citizenID
        record.Add "department", department
        record.Add "section", section
        record.Add "position", position

        record.Add "basicSalary", NumberByField_(wsSalary, rowIndex, salaryMap, "basicSalary")
        record.Add "workingDays", NumberByField_(wsSalary, rowIndex, salaryMap, "workingDays")
        record.Add "holidayDays", NumberByField_(wsSalary, rowIndex, salaryMap, "holidayDays")
        record.Add "paidLeaveDays", NumberByField_(wsSalary, rowIndex, salaryMap, "paidLeaveDays")
        record.Add "unpaidLeaveDays", NumberByField_(wsSalary, rowIndex, salaryMap, "unpaidLeaveDays")
        record.Add "overtimeHours", NumberByField_(wsSalary, rowIndex, salaryMap, "overtimeHours")
        record.Add "holidayWorkHours", NumberByField_(wsSalary, rowIndex, salaryMap, "holidayWorkHours")
        record.Add "holidayOvertimeHours", NumberByField_(wsSalary, rowIndex, salaryMap, "holidayOvertimeHours")
        record.Add "minimumRegionalLeaveDays", NumberByField_(wsSalary, rowIndex, salaryMap, "minimumRegionalLeaveDays")
        record.Add "nightHolidayOvertimeHours", NumberByField_(wsSalary, rowIndex, salaryMap, "nightHolidayOvertimeHours")
        record.Add "nightOvertimeHours", NumberByField_(wsSalary, rowIndex, salaryMap, "nightOvertimeHours")
        record.Add "holidayWorkDayHours", NumberByField_(wsSalary, rowIndex, salaryMap, "holidayWorkDayHours")
        record.Add "holidayOvertimeDayHours", NumberByField_(wsSalary, rowIndex, salaryMap, "holidayOvertimeDayHours")
        record.Add "nightHolidayOvertimeDayHours", NumberByField_(wsSalary, rowIndex, salaryMap, "nightHolidayOvertimeDayHours")
        record.Add "nightShiftDays", NumberByField_(wsSalary, rowIndex, salaryMap, "nightShiftDays")

        record.Add "otherMoney", NumberByField_(wsSalary, rowIndex, salaryMap, "otherMoney")
        record.Add "disciplinaryMoney", NumberByField_(wsSalary, rowIndex, salaryMap, "disciplinaryMoney")
        record.Add "loyalty2Years", NumberByField_(wsSalary, rowIndex, salaryMap, "loyalty2Years")
        record.Add "loyalty5Years", NumberByField_(wsSalary, rowIndex, salaryMap, "loyalty5Years")
        record.Add "loyalty10Years", NumberByField_(wsSalary, rowIndex, salaryMap, "loyalty10Years")
        record.Add "housingAllowance", NumberByField_(wsSalary, rowIndex, salaryMap, "housingAllowance")
        record.Add "transportationAllowance", NumberByField_(wsSalary, rowIndex, salaryMap, "transportationAllowance")
        record.Add "attendanceBonus", NumberByField_(wsSalary, rowIndex, salaryMap, "attendanceBonus")
        record.Add "severanceAndUnusedLeave", NumberByField_(wsSalary, rowIndex, salaryMap, "severanceAndUnusedLeave")

        record.Add "monthlySalary", NumberByField_(wsSalary, rowIndex, salaryMap, "monthlySalary")
        record.Add "overtimeSalary", NumberByField_(wsSalary, rowIndex, salaryMap, "overtimeSalary")
        record.Add "commissionAndOverTargetBonus", NumberByField_(wsSalary, rowIndex, salaryMap, "commissionAndOverTargetBonus")

        record.Add "otherIncome", OtherIncomeTotal_(wsSalary, rowIndex, salaryMap)

        record.Add "otherDeductions", NumberByField_(wsSalary, rowIndex, salaryMap, "otherDeductions")
        record.Add "advancePayment", NumberByField_(wsSalary, rowIndex, salaryMap, "advancePayment")
        record.Add "socialInsurance", NumberByField_(wsSalary, rowIndex, salaryMap, "socialInsurance")
        record.Add "healthInsurance", NumberByField_(wsSalary, rowIndex, salaryMap, "healthInsurance")
        record.Add "unemploymentInsurance", NumberByField_(wsSalary, rowIndex, salaryMap, "unemploymentInsurance")
        record.Add "personalIncomeTax", NumberByField_(wsSalary, rowIndex, salaryMap, "personalIncomeTax")
        record.Add "grossIncome", NumberByField_(wsSalary, rowIndex, salaryMap, "grossIncome")
        record.Add "netSalary", NumberByField_(wsSalary, rowIndex, salaryMap, "netSalary")

        ' Keep non-core Excel columns available for future use.
        ' These are sent in ExtraData but are ignored by the current GAS backend
        ' until a matching PayrollConfig field is configured there.
        record.Add "extraData", BuildExtraData_(wsSalary, rowIndex, extraHeaders)

        result.Add record

ContinueSalaryRow:
    Next rowIndex

    wb.Close SaveChanges:=False
    Set wb = Nothing

    LogMessage "[" & factoryName & "] Doc duoc " & CStr(result.Count) & " dong."
    Set ReadPayrollWorkbook_ = result
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Err.Raise Err.Number, "ReadPayrollWorkbook_", Err.Description
End Function

Private Function FindPayrollDataLastRow_(ByVal ws As Worksheet, ByVal nameColumn As Long) As Long
    Dim lastRow As Long
    Dim codeCol As Long
    Dim r As Long
    Dim employeeName As String
    Dim employeeCode As String
    Dim lastDataRow As Long

    ' Do not stop at the first TOTAL/SUBTOTAL row.
    ' Real payroll files can contain multiple TOTAL rows inside the sheet.
    lastRow = LastUsedRow_(ws, nameColumn)

    If lastRow < SHEET_START_ROW Then
        FindPayrollDataLastRow_ = SHEET_START_ROW - 1
        Exit Function
    End If

    codeCol = FindHeaderColumn_(ws, Array("Employee Code", "EmployeeCode", "Emp Code", "Mã nhân viên", "Ma nhan vien", "Mã NV"), SHEET_START_ROW - 1)
    If codeCol = 0 Then codeCol = 4

    lastDataRow = SHEET_START_ROW - 1

    For r = SHEET_START_ROW To lastRow
        employeeName = CellText_(ws.Cells(r, nameColumn))
        employeeCode = CellText_(ws.Cells(r, codeCol))

        ' A valid employee data row must have both employee code and name.
        If Len(employeeName) > 0 And Len(employeeCode) > 0 Then
            If Not IsSummaryRow_(employeeName, employeeCode) Then
                lastDataRow = r
            End If
        End If
    Next r

    FindPayrollDataLastRow_ = lastDataRow
End Function

Private Function IsSummaryRow_(ByVal employeeName As String, ByVal employeeCode As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(employeeName))
    s = Replace$(s, " ", "")
    s = Replace$(s, vbTab, "")

    If Left$(s, 5) = "TOTAL" Or Left$(s, 7) = "SUBTOTAL" Then
        IsSummaryRow_ = True
        Exit Function
    End If

    If InStr(1, s, "TỔNG", vbTextCompare) > 0 Or _
       InStr(1, s, "TONG", vbTextCompare) > 0 Then
        IsSummaryRow_ = True
        Exit Function
    End If

    IsSummaryRow_ = False
End Function

Private Function BuildDSCNVMap_(ByVal ws As Worksheet) As Object

    Dim map As Object
    Dim headerMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim employeeName As String
    Dim employeeCode As String
    Dim key As String
    Dim item As Object

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    Set headerMap = BuildDSCNVHeaderMap_(ws)
    lastRow = LastUsedRow_(ws, headerMap("name"))

    For rowIndex = SHEET_START_ROW To lastRow
        employeeName = CellText_(ws.Cells(rowIndex, headerMap("name")))
        employeeCode = ""

        If CLng(headerMap("employeeCode")) > 0 Then
            employeeCode = CellText_(ws.Cells(rowIndex, headerMap("employeeCode")))
        End If

        If Len(employeeName) > 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            item.CompareMode = vbTextCompare

            item.Add "EmployeeCode", employeeCode
            item.Add "CitizenID", ReadCitizenID_(ws.Cells(rowIndex, headerMap("citizenID")))
            item.Add "Department", CellText_(ws.Cells(rowIndex, headerMap("department")))
            item.Add "Section", CellText_(ws.Cells(rowIndex, headerMap("section")))
            item.Add "Position", CellText_(ws.Cells(rowIndex, headerMap("position")))

            ' Prefer employee-code index when available.
            If Len(employeeCode) > 0 Then
                key = "CODE|" & UCase$(employeeCode)
                If Not map.Exists(key) Then
                    map.Add key, item
                Else
                    LogMessage "CANH BAO: Trung ma nhan vien trong DSCNV: " & employeeCode
                End If
            End If

            ' Also maintain a name index as a fallback for legacy files.
            key = "NAME|" & NormalizeName_(employeeName)
            If Not map.Exists(key) Then
                map.Add key, item
            Else
                LogMessage "CANH BAO: Trung ten trong DSCNV: " & employeeName
            End If
        End If
    Next rowIndex

    Set BuildDSCNVMap_ = map
End Function

Private Function BuildSalaryHeaderMap_(ByVal ws As Worksheet) As Object
    Dim map As Object
    Dim specs As Collection
    Dim i As Long
    Dim col As Long
    Dim key As String
    Dim required As Boolean
    Dim legacyCol As Long

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    Set specs = SalaryFieldSpecs_()

    For i = 1 To specs.Count
        key = CStr(specs(i)(0))
        required = CBool(specs(i)(2))
        legacyCol = CLng(specs(i)(3))
        col = FindHeaderColumn_(ws, specs(i)(1), SHEET_START_ROW - 1)

        If col = 0 And legacyCol > 0 Then
            col = legacyCol
            LogMessage "CANH BAO: Khong tim thay Header cho '" & key & "'. Dung cot legacy " & CStr(legacyCol) & "."
        End If

        If col = 0 And required Then
            Err.Raise vbObjectError + 2101, "BuildSalaryHeaderMap_", _
                      "Khong tim thay cot bat buoc '" & key & "'. Alias: " & Join(specs(i)(1), ", ")
        End If

        map.Add key, col
    Next i

    LogHeaderMap_ ws, map, specs
    Set BuildSalaryHeaderMap_ = map
End Function

Private Function SalaryFieldSpecs_() As Collection
    Dim specs As Collection
    Set specs = New Collection

    ' key, aliases(), required, legacy fallback column
    SalarySpecAdd_ specs, "employeeCode", Array("Employee Code", "EmployeeCode", "Emp Code", "Mã nhân viên", "Ma nhan vien", "Mã NV"), True, 4
    SalarySpecAdd_ specs, "fullName", Array("Full Name", "FullName", "Employee Name", "Name", "Họ tên", "Ho ten", "Tên nhân viên", "Ten nhan vien"), True, 84
    SalarySpecAdd_ specs, "basicSalary", Array("Basic Salary", "BasicSalary", "Lương cơ bản", "Luong co ban"), False, 6
    SalarySpecAdd_ specs, "workingDays", Array("Working Days", "WorkingDays", "Ngày công", "Ngay cong", "Paid Working Days"), False, 8
    SalarySpecAdd_ specs, "holidayDays", Array("Holiday Days", "HolidayDays", "Ngày lễ", "Ngay le"), False, 9
    SalarySpecAdd_ specs, "paidLeaveDays", Array("Paid Leave Days", "PaidLeaveDays", "Phép hưởng lương", "Phep huong luong"), False, 10
    SalarySpecAdd_ specs, "unpaidLeaveDays", Array("Unpaid Leave Days", "UnpaidLeaveDays", "Phép không lương", "Phep khong luong"), False, 11
    SalarySpecAdd_ specs, "overtimeHours", Array("Overtime Hours", "OvertimeHours", "Giờ tăng ca", "Gio tang ca", "OT Hours"), False, 12
    SalarySpecAdd_ specs, "holidayWorkHours", Array("Holiday Work Hours", "HolidayWorkHours", "Giờ làm ngày lễ", "Gio lam ngay le"), False, 13
    SalarySpecAdd_ specs, "holidayOvertimeHours", Array("Holiday Overtime Hours", "HolidayOvertimeHours", "Giờ OT ngày lễ", "Gio OT ngay le"), False, 14
    SalarySpecAdd_ specs, "minimumRegionalLeaveDays", Array("Minimum Regional Leave Days", "MinimumRegionalLeaveDays", "Ngày phép tối thiểu vùng", "Ngay phep toi thieu vung"), False, 15
    SalarySpecAdd_ specs, "nightHolidayOvertimeHours", Array("Night Holiday OT", "Night Holiday Overtime Hours", "NightHolidayOvertimeHours", "Giờ OT đêm ngày lễ", "Gio OT dem ngay le"), False, 16
    SalarySpecAdd_ specs, "nightOvertimeHours", Array("Night OT", "Night Overtime Hours", "NightOvertimeHours", "Giờ OT ban đêm", "Gio OT ban dem"), False, 17
    SalarySpecAdd_ specs, "holidayWorkDayHours", Array("Holiday Work Day Hours", "HolidayWorkDayHours", "Giờ ngày làm lễ", "Gio ngay lam le"), False, 18
    SalarySpecAdd_ specs, "holidayOvertimeDayHours", Array("Holiday OT Day Hours", "HolidayOvertimeDayHours", "Giờ ngày OT lễ", "Gio ngay OT le"), False, 19
    SalarySpecAdd_ specs, "nightHolidayOvertimeDayHours", Array("Night Holiday OT Day", "NightHolidayOvertimeDayHours", "Giờ ngày OT đêm lễ", "Gio ngay OT dem le"), False, 20
    SalarySpecAdd_ specs, "nightShiftDays", Array("Night Shift Days", "NightShiftDays", "Ngày ca đêm", "Ngay ca dem"), False, 21
    SalarySpecAdd_ specs, "otherMoney", Array("Other Money", "OtherMoney", "Tiền khác", "Tien khac"), False, 24
    SalarySpecAdd_ specs, "disciplinaryMoney", Array("Disciplinary Money", "DisciplinaryMoney", "Tiền kỷ luật", "Tien ky luat"), False, 25
    SalarySpecAdd_ specs, "loyalty2Years", Array("Loyalty 2 Years", "Loyalty2Years", "Thâm niên 2 năm", "Tham nien 2 nam"), False, 26
    SalarySpecAdd_ specs, "loyalty5Years", Array("Loyalty 5 Years", "Loyalty5Years", "Thâm niên 5 năm", "Tham nien 5 nam"), False, 27
    SalarySpecAdd_ specs, "loyalty10Years", Array("Loyalty 10 Years", "Loyalty10Years", "Thâm niên 10 năm", "Tham nien 10 nam"), False, 28
    SalarySpecAdd_ specs, "housingAllowance", Array("Housing", "Housing Allowance", "HousingAllowance", "Phụ cấp nhà ở", "Phu cap nha o"), False, 29
    SalarySpecAdd_ specs, "transportationAllowance", Array("Transport", "Transportation Allowance", "TransportationAllowance", "Phụ cấp đi lại", "Phu cap di lai"), False, 30
    SalarySpecAdd_ specs, "attendanceBonus", Array("Attendance", "Attendance Bonus", "AttendanceBonus", "Chuyên cần", "Chuyen can"), False, 31
    SalarySpecAdd_ specs, "severanceAndUnusedLeave", Array("Severance Unused Leave", "SeveranceAndUnusedLeave", "Trợ cấp thôi việc phép tồn", "Tro cap thoi viec phep ton"), False, 32
    SalarySpecAdd_ specs, "socialInsurance", Array("Social Insurance", "SocialInsurance", "BHXH", "BHXH Employee", "Bảo hiểm xã hội"), False, 33
    SalarySpecAdd_ specs, "healthInsurance", Array("Health Insurance", "HealthInsurance", "BHYT", "BHYT Employee", "Bảo hiểm y tế"), False, 34
    SalarySpecAdd_ specs, "unemploymentInsurance", Array("Unemployment Insurance", "UnemploymentInsurance", "BHTN", "BHTN Employee", "Bảo hiểm thất nghiệp"), False, 35
    SalarySpecAdd_ specs, "otherDeductions", Array("Other Deductions", "OtherDeductions", "Khấu trừ khác", "Khau tru khac"), False, 36
    SalarySpecAdd_ specs, "advancePayment", Array("Advance", "Advance Payment", "AdvancePayment", "Tạm ứng", "Tam ung"), False, 37
    SalarySpecAdd_ specs, "monthlySalary", Array("Monthly Salary", "MonthlySalary", "Lương tháng", "Luong thang"), False, 40
    SalarySpecAdd_ specs, "overtimeSalary", Array("Overtime Salary", "OvertimeSalary", "Tiền tăng ca", "Tien tang ca"), False, 41
    SalarySpecAdd_ specs, "commissionAndOverTargetBonus", Array("Commission", "Commission And Over Target Bonus", "CommissionAndOverTargetBonus", "Hoa hồng", "Hoa hong"), False, 43
    SalarySpecAdd_ specs, "grossIncome", Array("Gross Income", "GrossIncome", "Tổng thu nhập", "Tong thu nhap"), False, 44
    SalarySpecAdd_ specs, "personalIncomeTax", Array("Personal Income Tax", "PersonalIncomeTax", "PIT", "Thuế TNCN", "Thue TNCN"), False, 46
    SalarySpecAdd_ specs, "netSalary", Array("Net Salary", "NetSalary", "Thực nhận", "Thuc nhan", "Net Pay"), False, 49

    Set SalaryFieldSpecs_ = specs
End Function

Private Sub SalarySpecAdd_(ByVal specs As Collection, ByVal key As String, ByVal aliases As Variant, ByVal required As Boolean, ByVal legacyCol As Long)
    specs.Add Array(key, aliases, required, legacyCol)
End Sub

Private Function BuildDSCNVHeaderMap_(ByVal ws As Worksheet) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    ' Employee code is optional because some legacy DSCNV files do not contain it.
    map.Add "employeeCode", FindHeaderColumn_(ws, Array("Employee Code", "EmployeeCode", "Emp Code", "Mã nhân viên", "Ma nhan vien", "Mã NV"), SHEET_START_ROW - 1)
    map.Add "name", FindHeaderWithFallback_(ws, Array("Name", "Full Name", "Employee Name", "Họ tên", "Ho ten", "Tên nhân viên"), 2, "DSCNV Name")
    map.Add "citizenID", FindHeaderWithFallback_(ws, Array("Citizen ID", "CitizenID", "CCCD", "CMND", "ID Card", "Số CCCD", "So CCCD"), 8, "DSCNV CitizenID")
    map.Add "department", FindHeaderWithFallback_(ws, Array("Department", "Dept", "Bộ phận", "Bo phan"), 32, "DSCNV Department")
    map.Add "section", FindHeaderWithFallback_(ws, Array("Section", "Team", "Tổ", "To"), 33, "DSCNV Section")
    map.Add "position", FindHeaderWithFallback_(ws, Array("Position", "Job Title", "Chức vụ", "Chuc vu"), 34, "DSCNV Position")

    Set BuildDSCNVHeaderMap_ = map
End Function

Private Function FindHeaderWithFallback_(ByVal ws As Worksheet, ByVal aliases As Variant, ByVal legacyCol As Long, ByVal label As String) As Long
    Dim col As Long
    col = FindHeaderColumn_(ws, aliases, SHEET_START_ROW - 1)
    If col = 0 Then
        col = legacyCol
        LogMessage "CANH BAO: Khong tim thay Header cho " & label & ". Dung cot legacy " & CStr(legacyCol) & "."
    End If
    FindHeaderWithFallback_ = col
End Function

Private Function FindHeaderColumn_(ByVal ws As Worksheet, ByVal aliases As Variant, ByVal maxHeaderRow As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    Dim normalizedCell As String, aliasItem As Variant
    Dim target As String

    For r = 1 To maxHeaderRow
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            normalizedCell = NormalizeHeader_(CStr(ws.Cells(r, c).Value2))
            If Len(normalizedCell) > 0 Then
                For Each aliasItem In aliases
                    target = NormalizeHeader_(CStr(aliasItem))
                    If normalizedCell = target Then
                        FindHeaderColumn_ = c
                        Exit Function
                    End If
                Next aliasItem
            End If
        Next c
    Next r

    FindHeaderColumn_ = 0
End Function

Private Function NormalizeHeader_(ByVal value As String) As String
    Dim s As String
    s = LCase$(Trim$(value))
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, "_", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, ":", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeHeader_ = s
End Function

Private Function OtherIncomeTotal_(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal salaryMap As Object) As Double
    Dim total As Double

    total = 0
    total = total + NumberByField_(ws, rowIndex, salaryMap, "otherMoney")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "disciplinaryMoney")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "loyalty2Years")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "loyalty5Years")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "loyalty10Years")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "housingAllowance")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "transportationAllowance")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "attendanceBonus")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "commissionAndOverTargetBonus")
    total = total + NumberByField_(ws, rowIndex, salaryMap, "severanceAndUnusedLeave")

    OtherIncomeTotal_ = total
End Function

Private Function NumberByField_(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal headerMap As Object, ByVal fieldKey As String) As Double
    If headerMap.Exists(fieldKey) Then
        If CLng(headerMap(fieldKey)) > 0 Then
            NumberByField_ = CellNumber_(ws.Cells(rowIndex, CLng(headerMap(fieldKey))))
            Exit Function
        End If
    End If
    NumberByField_ = 0
End Function

Private Function CellText_(ByVal cell As Range) As String
    On Error GoTo SafeEmpty
    CellText_ = Trim$(CStr(cell.Value2))
    Exit Function
SafeEmpty:
    CellText_ = ""
End Function

Private Function CollectExtraHeaders_(ByVal ws As Worksheet, ByVal salaryMap As Object) As Object
    Dim result As Object
    Dim used As Object
    Dim r As Long, c As Long, lastCol As Long
    Dim header As String, key As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In salaryMap.Keys
        If CLng(salaryMap(k)) > 0 Then used(CStr(salaryMap(k))) = True
    Next k

    For r = 1 To SHEET_START_ROW - 1
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            header = Trim$(CStr(ws.Cells(r, c).Value2))
            If Len(header) > 0 Then
                If Not used.Exists(CStr(c)) Then
                    key = SafeFieldKey_(header)
                    If Len(key) > 0 Then
                        If Not result.Exists(key) Then result.Add key, c
                    End If
                End If
            End If
        Next c
    Next r

    Set CollectExtraHeaders_ = result
End Function

Private Function BuildExtraData_(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal extraHeaders As Object) As Object
    Dim result As Object
    Dim k As Variant
    Dim value As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For Each k In extraHeaders.Keys
        value = ws.Cells(rowIndex, CLng(extraHeaders(k))).Value2
        If Len(Trim$(CStr(value))) > 0 Then
            If IsNumeric(value) Then
                result.Add CStr(k), CDbl(value)
            Else
                result.Add CStr(k), CStr(value)
            End If
        End If
    Next k

    Set BuildExtraData_ = result
End Function

Private Function SafeFieldKey_(ByVal header As String) As String
    Dim s As String, i As Long, ch As String, code As Long, out As String
    s = LCase$(Trim$(header))
    s = Replace(s, "đ", "d")
    s = Replace(s, "Đ", "d")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 97 And code <= 122) Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i
    Do While InStr(out, "__") > 0
        out = Replace(out, "__", "_")
    Loop
    out = Trim$(out)
    If Left$(out, 1) = "_" Then out = Mid$(out, 2)
    If Right$(out, 1) = "_" Then out = Left$(out, Len(out) - 1)
    SafeFieldKey_ = out
End Function

Private Sub LogHeaderMap_(ByVal ws As Worksheet, ByVal map As Object, ByVal specs As Collection)
    Dim i As Long, key As String, col As Long
    LogMessage "--- HEADER MAPPING: " & ws.Name & " ---"
    For i = 1 To specs.Count
        key = CStr(specs(i)(0))
        col = CLng(map(key))
        If col > 0 Then
            LogMessage "MAP " & key & " => col " & CStr(col) & " | Header='" & CStr(ws.Cells(FindHeaderRow_(ws, col), col).Value2) & "'"
        Else
            LogMessage "MAP " & key & " => NOT FOUND / optional"
        End If
    Next i
End Sub

Private Function FindHeaderRow_(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    For r = 1 To SHEET_START_ROW - 1
        If Len(Trim$(CStr(ws.Cells(r, col).Value2))) > 0 Then
            FindHeaderRow_ = r
            Exit Function
        End If
    Next r
    FindHeaderRow_ = SHEET_START_ROW - 1
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
                    CStr(cell.text)

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
        Trim$(CStr(cell.text))

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
            ws.Rows.count, _
            columnNumber).End(xlUp).row

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

    For i = 1 To records.count

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
            "grossIncome", _
        record("grossIncome"), True
            
    AddJsonNumber_ json, _
        "netSalary", _
        record("netSalary"), True

    AddJsonNumber_ json, _
        "nightHolidayOvertimeDayHours", _
        record("nightHolidayOvertimeDayHours"), True

    AddJsonObject_ json, _
        "extraData", _
        record("extraData"), False

    json = json & "}"

    RecordToJson_ = json

End Function

Private Sub AddJsonObject_(ByRef json As String, ByVal key As String, ByVal value As Object, ByVal addComma As Boolean)
    Dim k As Variant
    Dim first As Boolean
    Dim v As Variant

    json = json & """" & key & """:{"
    first = True
    If Not value Is Nothing Then
        For Each k In value.Keys
            If Not first Then json = json & ","
            v = value(k)
            json = json & """" & JsonEscape_(CStr(k)) & """:"
            If IsNumeric(v) And VarType(v) <> vbString Then
                json = json & JsonNumberString_(CDbl(v))
            Else
                json = json & """" & JsonEscape_(CStr(v)) & """"
            End If
            first = False
        Next k
    End If
    json = json & "}"
    If addComma Then json = json & ","
End Sub

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
        """" & key & """:""" & _
        JsonNumberString_(value) & """"

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
        CStr(http.responseText)

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
    stream.position = 0
    stream.Type = 1

    bytes = stream.Read
    length = UBound(bytes) - LBound(bytes) + 1

    ' ADODB.Stream writes UTF-8 BOM (EF BB BF).
    If length >= 3 Then
        If bytes(0) = &HEF And _
           bytes(1) = &HBB And _
           bytes(2) = &HBF Then

            stream.position = 3
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
        ThisWorkbook.path & _
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

