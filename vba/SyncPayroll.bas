Attribute VB_Name = "SyncPayroll"
Option Explicit

' =====================================================================
' ĐỒNG BỘ LƯƠNG XLSB -> GAS
' Chạy RunSync. Macro tự lấy THÁNG TRƯỚC tháng hiện tại.
' Ví dụ 14/09/2026 -> 08-2026.
' =====================================================================

Public Sub RunSync()
    Dim targetMonth As Integer, targetYear As Integer
    GetTargetMonthYear targetMonth, targetYear
    RunSyncForMonth targetMonth, targetYear
End Sub

' Dùng để test thủ công, ví dụ:
'   RunSyncForSpecificMonth
' rồi sửa TARGET_MONTH/TARGET_YEAR bên dưới nếu cần.
Public Sub RunSyncForSpecificMonth()
    Const TARGET_MONTH As Integer = 8
    Const TARGET_YEAR As Integer = 2026
    RunSyncForMonth TARGET_MONTH, TARGET_YEAR
End Sub

Private Sub RunSyncForMonth(ByVal targetMonth As Integer, ByVal targetYear As Integer)
    Dim thangStr As String
    thangStr = Format(targetMonth, "00") & "-" & CStr(targetYear)

    LogMessage "============================================================"
    LogMessage "BAT DAU DONG BO " & thangStr
    LogMessage "GAS_URL = " & GAS_URL

    Dim oldScreen As Boolean, oldAlerts As Boolean
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo FatalError
    SyncOneWorkshop SNACK_ROOT, SNACK_FILE_PREFIX, SNACK_LABEL, thangStr
    SyncOneWorkshop FLEXIBLE_ROOT, FLEXIBLE_FILE_PREFIX, FLEXIBLE_LABEL, thangStr

    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    LogMessage "HOAN TAT DONG BO " & thangStr
    MsgBox "Đồng bộ hoàn tất. Xem log tại:" & vbCrLf & Environ$("TEMP") & "\SyncPayroll.log", vbInformation, "Sync Payroll"
    Exit Sub

FatalError:
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    LogMessage "LOI NGHIEM TRONG: " & Err.Number & " - " & Err.Description
    MsgBox "Đồng bộ gặp lỗi:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Log: " & Environ$("TEMP") & "\SyncPayroll.log", vbCritical, "Sync Payroll"
End Sub

Private Sub GetTargetMonthYear(ByRef outMonth As Integer, ByRef outYear As Integer)
    Dim d As Date
    d = DateAdd("m", -1, Date)
    outMonth = Month(d)
    outYear = Year(d)
End Sub

Private Sub SyncOneWorkshop(ByVal rootPath As String, ByVal filePrefix As String, _
                            ByVal label As String, ByVal thangStr As String)
    Dim filePath As String
    filePath = rootPath & "\" & Right$(thangStr, 4) & "\" & filePrefix & " " & thangStr & ".xlsb"
    LogMessage "[" & label & "] FILE = " & filePath

    If Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then
        LogMessage "[" & label & "] KHONG TIM THAY FILE. Kiem tra UNC/network/path."
        Exit Sub
    End If

    Dim wb As Workbook
    On Error GoTo OpenFailed
    Set wb = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True, _
                            Password:=XLSB_PASSWORD, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    On Error GoTo DataFailed
    LogMessage "[" & label & "] Mo file OK."

    Dim wsSalary As Worksheet, wsDscnv As Worksheet
    Set wsSalary = GetSheetOrNothing(wb, SHEET_SALARY)
    Set wsDscnv = GetSheetOrNothing(wb, SHEET_DSCNV)
    If wsSalary Is Nothing Then Err.Raise vbObjectError + 100, , "Khong tim thay sheet Salary."
    If wsDscnv Is Nothing Then Err.Raise vbObjectError + 101, , "Khong tim thay sheet DSCNV."

    LogMessage "[" & label & "] Salary rows last = " & wsSalary.Cells(wsSalary.Rows.Count, ColLetterToNumber(COL_SALARY_MANV)).End(xlUp).Row
    LogMessage "[" & label & "] DSCNV rows last = " & wsDscnv.Cells(wsDscnv.Rows.Count, ColLetterToNumber(COL_DSCNV_HOTEN)).End(xlUp).Row

    Dim lookup As Object, employees As Collection, payroll As Collection
    Set lookup = BuildDscnvLookup(wsDscnv)
    Set employees = New Collection
    Set payroll = New Collection
    ReadSalarySheet wsSalary, lookup, employees, payroll

    wb.Close SaveChanges:=False
    Set wb = Nothing

    LogMessage "[" & label & "] employees=" & employees.Count & "; payroll=" & payroll.Count
    If payroll.Count = 0 Then
        LogMessage "[" & label & "] KHONG CO DONG LUONG. Khong gui GAS."
        Exit Sub
    End If

    Dim jsonBody As String, responseText As String
    jsonBody = "{" & _
        Json.JsonString("action") & ":" & Json.JsonString("sync") & "," & _
        Json.JsonString("apiKey") & ":" & Json.JsonString(SYNC_API_KEY) & "," & _
        Json.JsonString("xuong") & ":" & Json.JsonString(label) & "," & _
        Json.JsonString("thang") & ":" & Json.JsonString(thangStr) & "," & _
        Json.JsonString("clientStartedAt") & ":" & Json.JsonString(CStr(CDbl(Now))) & "," & _
        Json.JsonString("employees") & ":" & Json.CollectionToJsonArray(employees) & "," & _
        Json.JsonString("payroll") & ":" & Json.CollectionToJsonArray(payroll) & _
        "}"

    LogMessage "[" & label & "] JSON length = " & Len(jsonBody)
    responseText = PostJsonUtf8(GAS_URL, jsonBody)
    LogMessage "[" & label & "] SERVER RESPONSE = " & responseText
    Exit Sub

OpenFailed:
    LogMessage "[" & label & "] MO FILE THAT BAI: " & Err.Number & " - " & Err.Description
    Exit Sub

DataFailed:
    LogMessage "[" & label & "] LOI DOC DU LIEU: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function GetSheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(sheetName), vbTextCompare) = 0 Then
            Set GetSheetOrNothing = ws
            Exit Function
        End If
    Next ws
    Set GetSheetOrNothing = Nothing
End Function

Private Function BuildDscnvLookup(ByVal ws As Worksheet) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, ColLetterToNumber(COL_DSCNV_HOTEN)).End(xlUp).Row
    Dim r As Long, hoTen As String, key As String, info As Object
    For r = DATA_START_ROW To lastRow
        hoTen = SafeText(GetCellText(ws, r, COL_DSCNV_HOTEN))
        key = NormalizeName(hoTen)
        If key <> "" Then
            Set info = CreateObject("Scripting.Dictionary")
            info("cccd") = NormalizeCCCDText(GetCellText(ws, r, COL_DSCNV_CCCD))
            info("phongBan") = SafeText(GetCellText(ws, r, COL_DSCNV_PHONGBAN))
            info("boPhan") = SafeText(GetCellText(ws, r, COL_DSCNV_BOPHAN))
            info("chucVu") = SafeText(GetCellText(ws, r, COL_DSCNV_CHUCVU))
            If Not result.Exists(key) Then result.Add key, info
        End If
    Next r
    Set BuildDscnvLookup = result
End Function

Private Sub ReadSalarySheet(ByVal ws As Worksheet, ByVal dscnvLookup As Object, _
                            ByRef employees As Collection, ByRef payroll As Collection)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, ColLetterToNumber(COL_SALARY_MANV)).End(xlUp).Row
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    Dim r As Long, maNV As String, hoTen As String, key As String
    For r = DATA_START_ROW To lastRow
        maNV = SafeText(GetCellText(ws, r, COL_SALARY_MANV))
        If maNV <> "" Then
            hoTen = SafeText(GetCellText(ws, r, COL_SALARY_HOTEN))
            key = NormalizeName(hoTen)

            Dim cccd As String, phongBan As String, boPhan As String, chucVu As String
            cccd = "": phongBan = "": boPhan = "": chucVu = ""
            If dscnvLookup.Exists(key) Then
                Dim info As Object: Set info = dscnvLookup(key)
                cccd = info("cccd"): phongBan = info("phongBan"): boPhan = info("boPhan"): chucVu = info("chucVu")
            Else
                LogMessage "[CANH BAO] Khong match DSCNV: " & hoTen & " / " & maNV
            End If

            If Not seen.Exists(maNV) Then
                seen.Add maNV, True
                employees.Add "{" & _
                    Json.JsonString("maNV") & ":" & Json.JsonString(maNV) & "," & _
                    Json.JsonString("hoTen") & ":" & Json.JsonString(hoTen) & "," & _
                    Json.JsonString("cccd") & ":" & Json.JsonString(cccd) & "," & _
                    Json.JsonString("phongBan") & ":" & Json.JsonString(phongBan) & "," & _
                    Json.JsonString("boPhan") & ":" & Json.JsonString(boPhan) & "," & _
                    Json.JsonString("chucVu") & ":" & Json.JsonString(chucVu) & "}"
            End If

            Dim p As Object: Set p = CreateObject("Scripting.Dictionary")
            p("maNV") = maNV: p("hoTen") = hoTen
            p("luongCoBan") = GetCellNumber(ws,r,COL_SALARY_LUONGCOBAN)
            p("soNgayLamViec") = GetCellNumber(ws,r,COL_SALARY_SONGAYLAMVIEC)
            p("soNgayLe") = GetCellNumber(ws,r,COL_SALARY_SONGAYLE)
            p("soNgayNghiHuongLuong") = GetCellNumber(ws,r,COL_SALARY_SONGAYNGHIHUONGLUONG)
            p("soNgayNghiKhongLuong") = GetCellNumber(ws,r,COL_SALARY_SONGAYNGHIKHONGLUONG)
            p("soNgayNghiHuongLuongToiThieuVung") = GetCellNumber(ws,r,COL_SALARY_SONGAYNGHIHUONGLUONGTOITHIEUVUNG)
            p("luongThang") = GetCellNumber(ws,r,COL_SALARY_LUONGTHANG)
            p("soGioNgoaiGio") = GetCellNumber(ws,r,COL_SALARY_SOGIONGOAIGIO)
            p("soGioNgayNghi") = GetCellNumber(ws,r,COL_SALARY_SOGIONGAYNGHI)
            p("soGioNgoaiGioNgayNghi") = GetCellNumber(ws,r,COL_SALARY_SOGIONGOAIGIONGAYNGHI)
            p("soGioTangCaDem") = GetCellNumber(ws,r,COL_SALARY_SOGIOTANGCADEM)
            p("soGioLamNgayLe") = GetCellNumber(ws,r,COL_SALARY_SOGIOLAMNGAYLE)
            p("soGioNgoaiGioNgayLe") = GetCellNumber(ws,r,COL_SALARY_SOGIONGOAIGIONGAYLE)
            p("soGioTangCaDemNgayLe") = GetCellNumber(ws,r,COL_SALARY_SOGIOTANGCADEMNGAYLE)
            p("soNgayLamCaDem") = GetCellNumber(ws,r,COL_SALARY_SONGAYLAMCADEM)
            p("luongNgoaiGio") = GetCellNumber(ws,r,COL_SALARY_LUONGNGOAIGIO)
            p("tienKhac") = GetCellNumber(ws,r,COL_SALARY_TIENKHAC)
            p("tienKyLuat") = GetCellNumber(ws,r,COL_SALARY_TIENKYLUAT)
            p("tienGanBo2Nam") = GetCellNumber(ws,r,COL_SALARY_TIENGANBO2NAM)
            p("tienGanBo5Nam") = GetCellNumber(ws,r,COL_SALARY_TIENGANBO5NAM)
            p("tienGanBo10Nam") = GetCellNumber(ws,r,COL_SALARY_TIENGANBO10NAM)
            p("tienNhaO") = GetCellNumber(ws,r,COL_SALARY_TIENNHAO)
            p("tienDiLai") = GetCellNumber(ws,r,COL_SALARY_TIENDILAI)
            p("tienThuongChuyenCan") = GetCellNumber(ws,r,COL_SALARY_TIENTHUONGCHUYENCAN)
            p("hoaHongThuongVuotDinhMuc") = GetCellNumber(ws,r,COL_SALARY_HOAHONGTHUONGVUOTDINHMUC)
            p("troCapThoiViecPhepNam") = GetCellNumber(ws,r,COL_SALARY_TROCAPTHOIVIECPHEPNAM)
            p("tongKhoanThuNhap") = GetCellNumber(ws,r,COL_SALARY_TONGKHOANTHUNHAP)
            p("bhxh") = GetCellNumber(ws,r,COL_SALARY_BHXH)
            p("bhyt") = GetCellNumber(ws,r,COL_SALARY_BHYT)
            p("bhtn") = GetCellNumber(ws,r,COL_SALARY_BHTN)
            p("thueThuNhap") = GetCellNumber(ws,r,COL_SALARY_THUETHUNHAP)
            p("tamUng") = GetCellNumber(ws,r,COL_SALARY_TAMUNG)
            p("khauTruKhac") = GetCellNumber(ws,r,COL_SALARY_KHAUTRUKHAC)
            p("luongThucLinh") = GetCellNumber(ws,r,COL_SALARY_LUONGTHUCLINH)

            Dim nums As Collection: Set nums = New Collection
            Dim k As Variant
            For Each k In p.Keys
                If k <> "maNV" And k <> "hoTen" Then nums.Add CStr(k)
            Next k
            payroll.Add Json.DictToJson(p, nums)
        End If
    Next r
End Sub

Public Function ColLetterToNumber(ByVal letters As String) As Long
    Dim i As Long, n As Long, ch As Long, s As String
    s = UCase$(Trim$(letters))
    For i = 1 To Len(s)
        ch = Asc(Mid$(s,i,1)) - 64
        If ch < 1 Or ch > 26 Then Err.Raise vbObjectError + 200, , "Cot khong hop le: " & letters
        n = n * 26 + ch
    Next i
    ColLetterToNumber = n
End Function

Private Function GetCellText(ByVal ws As Worksheet, ByVal row As Long, ByVal colLetter As String) As String
    Dim c As Range: Set c = ws.Cells(row, ColLetterToNumber(colLetter))
    If IsError(c.Value) Or IsEmpty(c.Value) Then GetCellText = "": Exit Function
    If c.NumberFormat = "@" Then
        GetCellText = CStr(c.Value2)
    ElseIf VarType(c.Value2) = vbString Then
        GetCellText = CStr(c.Value2)
    Else
        GetCellText = CStr(c.Text)
        If Len(GetCellText) = 0 Or InStr(GetCellText, "###") > 0 Then GetCellText = CStr(c.Value2)
    End If
End Function

Private Function GetCellNumber(ByVal ws As Worksheet, ByVal row As Long, ByVal colLetter As String) As Double
    Dim v As Variant: v = ws.Cells(row, ColLetterToNumber(colLetter)).Value2
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then GetCellNumber = 0: Exit Function
    If IsNumeric(v) Then GetCellNumber = CDbl(v) Else GetCellNumber = 0
End Function

Private Function SafeText(ByVal s As String) As String
    s = Replace(s, Chr$(39), "")
    s = Replace(s, ChrW$(160), " ")
    SafeText = Trim$(s)
End Function

Private Function NormalizeName(ByVal s As String) As String
    s = SafeText(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeName = UCase$(s)
End Function

Private Function NormalizeCCCDText(ByVal s As String) As String
    s = SafeText(s)
    If Len(s) > 0 And Left$(s, 1) = "'" Then s = Mid$(s, 2)
    s = Replace(s, " ", "")
    ' CCCD phải là chuỗi số; không tự chèn số 0 vì chỉ Excel gốc mới biết số 0 có hay không.
    NormalizeCCCDText = s
End Function

Private Function PostJsonUtf8(ByVal url As String, ByVal jsonBody As String) As String
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim bodyBytes() As Byte: bodyBytes = Utf8BytesFromString(jsonBody)
    On Error GoTo HttpFailed
    http.Open "POST", url, False
    http.SetTimeouts 30000, 30000, 300000, 300000
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send bodyBytes
    LogMessage "HTTP Status = " & http.Status & " " & http.StatusText
    PostJsonUtf8 = http.ResponseText
    Exit Function
HttpFailed:
    PostJsonUtf8 = "LOI HTTP " & Err.Number & ": " & Err.Description
End Function

Private Function Utf8BytesFromString(ByVal s As String) As Byte()
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText s
    stream.Position = 0: stream.Type = 1: stream.Position = 3
    Utf8BytesFromString = stream.Read
    stream.Close
End Function

Public Sub LogMessage(ByVal msg As String)
    Debug.Print msg
    On Error Resume Next
    Dim path As String, f As Integer
    path = Environ$("TEMP") & "\SyncPayroll.log"
    f = FreeFile
    Open path For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    Close #f
    On Error GoTo 0
End Sub
