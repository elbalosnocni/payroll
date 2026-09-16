Attribute VB_Name = "SyncPayroll"
'-------------------------------------------------------------------------
' SyncPayroll.bas - PHIEN BAN ON DINH / NHANH
'
' Muc tieu:
'   - Chay 1 phat, khong nuot loi.
'   - Gui JSON UTF-8 dung Content-Type.
'   - Ho tro redirect cua Google Apps Script Web App.
'   - Giu retry cho loi mang / timeout / SYNC_BUSY / loi tam thoi.
'   - Giu nguyen cau truc payload, API key, du lieu luong va 2 xuong.
'   - Bao loi ro rang neu GAS tu choi request (FORBIDDEN, MISSING_FIELDS...).
'-------------------------------------------------------------------------
Option Explicit

Public Sub RunSync()
    Dim targetMonth As Integer, targetYear As Integer
    Dim thangStr As String
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim runError As String

    On Error GoTo EH

    GetTargetMonthYear targetMonth, targetYear
    thangStr = Format(targetMonth, "00") & "-" & CStr(targetYear)

    LogMessage "=== BAT DAU DONG BO LUONG THANG " & thangStr & " LUC " & Now & " ==="

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Chay lan luot, tranh 2 request GAS cung luc -> giam nguy co SYNC_BUSY.
    SyncOneWorkshop SNACK_ROOT, SNACK_FILE_PREFIX, SNACK_LABEL, targetMonth, targetYear, thangStr
    SyncOneWorkshop FLEXIBLE_ROOT, FLEXIBLE_FILE_PREFIX, FLEXIBLE_LABEL, targetMonth, targetYear, thangStr

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    If runError = "" Then
        LogMessage "=== KET THUC DONG BO LUC " & Now & " ==="
        MsgBox "Dong bo Payroll da hoan tat." & vbCrLf & _
               "Thang: " & thangStr & vbCrLf & vbCrLf & _
               "Chi tiet xem file log:" & vbCrLf & _
               Environ$("TEMP") & "\SyncPayroll.log", vbInformation, "Payroll Sync"
    Else
        LogMessage "=== DONG BO THAT BAI: " & runError & " ==="
        MsgBox "Dong bo Payroll that bai:" & vbCrLf & vbCrLf & _
               runError & vbCrLf & vbCrLf & _
               "Log: " & Environ$("TEMP") & "\SyncPayroll.log", _
               vbCritical, "Payroll Sync"
    End If
    Exit Sub

EH:
    runError = "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

' Thang can lay = thang truoc thang hien tai.
Private Sub GetTargetMonthYear(ByRef outMonth As Integer, ByRef outYear As Integer)
    Dim prevMonthDate As Date
    prevMonthDate = DateAdd("m", -1, Date)
    outMonth = Month(prevMonthDate)
    outYear = Year(prevMonthDate)
End Sub

Private Sub SyncOneWorkshop(ByVal rootPath As String, ByVal filePrefix As String, _
                            ByVal label As String, ByVal targetMonth As Integer, _
                            ByVal targetYear As Integer, ByVal thangStr As String)
    Dim filePath As String
    filePath = rootPath & "\" & CStr(targetYear) & "\" & filePrefix & " " & thangStr & ".xlsb"

    LogMessage "[" & label & "] Dang tim file: " & filePath

    If Dir(filePath) = "" Then
        Err.Raise vbObjectError + 2101, "SyncOneWorkshop", _
                  "[" & label & "] Khong tim thay file: " & filePath
    End If

    Dim wb As Workbook
    Dim wsSalary As Worksheet, wsDscnv As Worksheet
    Dim dscnvLookup As Object
    Dim employees As Collection, payroll As Collection
    Dim jsonBody As String, responseText As String

    On Error GoTo OpenFailed
    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=True, _
                            Password:=XLSB_PASSWORD, IgnoreReadOnlyRecommended:=True)
    On Error GoTo EH

    Set wsSalary = GetSheetOrNothing(wb, SHEET_SALARY)
    Set wsDscnv = GetSheetOrNothing(wb, SHEET_DSCNV)

    If wsSalary Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Err.Raise vbObjectError + 2102, "SyncOneWorkshop", _
                  "[" & label & "] Khong tim thay sheet '" & SHEET_SALARY & "'."
    End If

    If wsDscnv Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Err.Raise vbObjectError + 2103, "SyncOneWorkshop", _
                  "[" & label & "] Khong tim thay sheet '" & SHEET_DSCNV & "'."
    End If

    Set dscnvLookup = BuildDscnvLookup(wsDscnv)

    Set employees = New Collection
    Set payroll = New Collection
    ReadSalarySheet wsSalary, dscnvLookup, employees, payroll

    wb.Close SaveChanges:=False
    Set wb = Nothing

    LogMessage "[" & label & "] Doc duoc " & payroll.Count & _
               " dong luong, " & employees.Count & " nhan vien."

    If payroll.Count = 0 Then
        Err.Raise vbObjectError + 2104, "SyncOneWorkshop", _
                  "[" & label & "] Khong co du lieu luong de dong bo."
    End If

    jsonBody = "{" & _
        Json.JsonString("action") & ":" & Json.JsonString("sync") & "," & _
        Json.JsonString("apiKey") & ":" & Json.JsonString(SYNC_API_KEY) & "," & _
        Json.JsonString("xuong") & ":" & Json.JsonString(label) & "," & _
        Json.JsonString("thang") & ":" & Json.JsonString(thangStr) & "," & _
        Json.JsonString("employees") & ":" & Json.CollectionToJsonArray(employees) & "," & _
        Json.JsonString("payroll") & ":" & Json.CollectionToJsonArray(payroll) & _
        "}"

    LogMessage "[" & label & "] Dang gui " & Len(jsonBody) & " ky tu JSON len GAS..."
    responseText = PostJsonUtf8WithRetry(GAS_URL, jsonBody, label)
    LogMessage "[" & label & "] GAS OK: " & responseText
    Exit Sub

OpenFailed:
    Dim openErr As String
    openErr = "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 2105, "SyncOneWorkshop", _
              "[" & label & "] Loi khi mo file .xlsb (kiem tra duong dan/mat khau): " & openErr

EH:
    Dim eNo As Long, eDesc As String
    eNo = Err.Number
    eDesc = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise eNo, "SyncOneWorkshop", eDesc
End Sub

Private Function GetSheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' Doc sheet DSCNV, tra ve Dictionary: key = Ho ten (trim, chuan hoa), value =
' Dictionary(cccd, phongban, bophan, chucvu)
Private Function BuildDscnvLookup(ByVal ws As Worksheet) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToNumber(COL_DSCNV_HOTEN)).End(xlUp).Row

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim hoTen As String
        hoTen = SafeText(GetCellText(ws, r, COL_DSCNV_HOTEN))
        If hoTen <> "" Then
            Dim info As Object
            Set info = CreateObject("Scripting.Dictionary")
            info("cccd") = SafeText(GetCellText(ws, r, COL_DSCNV_CCCD))
            info("phongBan") = SafeText(GetCellText(ws, r, COL_DSCNV_PHONGBAN))
            info("boPhan") = SafeText(GetCellText(ws, r, COL_DSCNV_BOPHAN))
            info("chucVu") = SafeText(GetCellText(ws, r, COL_DSCNV_CHUCVU))

            Dim key As String
            key = NormalizeName(hoTen)
            If Not result.Exists(key) Then result.Add key, info
        End If
    Next r

    Set BuildDscnvLookup = result
End Function

' Doc sheet Salary, ghep DSCNV, dong goi employees + payroll.
Private Sub ReadSalarySheet(ByVal ws As Worksheet, ByVal dscnvLookup As Object, _
                            ByRef employees As Collection, ByRef payroll As Collection)
    Dim rawLastRow As Long
    rawLastRow = ws.Cells(ws.Rows.Count, ColLetterToNumber(COL_SALARY_MANV)).End(xlUp).Row

    Dim lastRow As Long
    lastRow = ExcludeTotalRow(ws, DATA_START_ROW, rawLastRow)

    Dim seenManv As Object
    Set seenManv = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim maNV As String
        maNV = SafeText(GetCellText(ws, r, COL_SALARY_MANV))
        If maNV = "" Then GoTo ContinueLoop

        Dim hoTen As String
        hoTen = SafeText(GetCellText(ws, r, COL_SALARY_HOTEN))

        Dim info As Object
        Dim cccd As String, phongBan As String, boPhan As String, chucVu As String
        cccd = "": phongBan = "": boPhan = "": chucVu = ""

        If dscnvLookup.Exists(NormalizeName(hoTen)) Then
            Set info = dscnvLookup(NormalizeName(hoTen))
            cccd = info("cccd")
            phongBan = info("phongBan")
            boPhan = info("boPhan")
            chucVu = info("chucVu")
        Else
            LogMessage "  [Canh bao] Khong khop DSCNV cho nhan vien: " & hoTen & _
                       " (Ma NV " & maNV & ")"
        End If

        If Not seenManv.Exists(maNV) Then
            seenManv.Add maNV, True

            Dim empJson As String
            empJson = "{" & _
                Json.JsonString("maNV") & ":" & Json.JsonString(maNV) & "," & _
                Json.JsonString("hoTen") & ":" & Json.JsonString(hoTen) & "," & _
                Json.JsonString("cccd") & ":" & Json.JsonString(cccd) & "," & _
                Json.JsonString("phongBan") & ":" & Json.JsonString(phongBan) & "," & _
                Json.JsonString("boPhan") & ":" & Json.JsonString(boPhan) & "," & _
                Json.JsonString("chucVu") & ":" & Json.JsonString(chucVu) & _
                "}"
            employees.Add empJson
        End If

        Dim p As Object
        Set p = CreateObject("Scripting.Dictionary")
        p("maNV") = maNV
        p("hoTen") = hoTen

        p("luongCoBan") = GetCellNumber(ws, r, COL_SALARY_LUONGCOBAN)
        p("soNgayLamViec") = GetCellNumber(ws, r, COL_SALARY_SONGAYLAMVIEC)
        p("soNgayLe") = GetCellNumber(ws, r, COL_SALARY_SONGAYLE)
        p("soNgayNghiHuongLuong") = GetCellNumber(ws, r, COL_SALARY_SONGAYNGHIHUONGLUONG)
        p("soNgayNghiKhongLuong") = GetCellNumber(ws, r, COL_SALARY_SONGAYNGHIKHONGLUONG)
        p("soNgayNghiHuongLuongToiThieuVung") = GetCellNumber(ws, r, COL_SALARY_SONGAYNGHIHUONGLUONGTOITHIEUVUNG)
        p("luongThang") = GetCellNumber(ws, r, COL_SALARY_LUONGTHANG)

        p("soGioNgoaiGio") = GetCellNumber(ws, r, COL_SALARY_SOGIONGOAIGIO)
        p("soGioNgayNghi") = GetCellNumber(ws, r, COL_SALARY_SOGIONGAYNGHI)
        p("soGioNgoaiGioNgayNghi") = GetCellNumber(ws, r, COL_SALARY_SOGIONGOAIGIONGAYNGHI)
        p("soGioTangCaDem") = GetCellNumber(ws, r, COL_SALARY_SOGIOTANGCADEM)
        p("soGioLamNgayLe") = GetCellNumber(ws, r, COL_SALARY_SOGIOLAMNGAYLE)
        p("soGioNgoaiGioNgayLe") = GetCellNumber(ws, r, COL_SALARY_SOGIONGOAIGIONGAYLE)
        p("soGioTangCaDemNgayLe") = GetCellNumber(ws, r, COL_SALARY_SOGIOTANGCADEMNGAYLE)
        p("soNgayLamCaDem") = GetCellNumber(ws, r, COL_SALARY_SONGAYLAMCADEM)
        p("luongNgoaiGio") = GetCellNumber(ws, r, COL_SALARY_LUONGNGOAIGIO)

        p("tienKhac") = GetCellNumber(ws, r, COL_SALARY_TIENKHAC)
        p("tienKyLuat") = GetCellNumber(ws, r, COL_SALARY_TIENKYLUAT)
        p("tienGanBo2Nam") = GetCellNumber(ws, r, COL_SALARY_TIENGANBO2NAM)
        p("tienGanBo5Nam") = GetCellNumber(ws, r, COL_SALARY_TIENGANBO5NAM)
        p("tienGanBo10Nam") = GetCellNumber(ws, r, COL_SALARY_TIENGANBO10NAM)
        p("tienNhaO") = GetCellNumber(ws, r, COL_SALARY_TIENNHAO)
        p("tienDiLai") = GetCellNumber(ws, r, COL_SALARY_TIENDILAI)
        p("tienThuongChuyenCan") = GetCellNumber(ws, r, COL_SALARY_TIENTHUONGCHUYENCAN)
        p("hoaHongThuongVuotDinhMuc") = GetCellNumber(ws, r, COL_SALARY_HOAHONGTHUONGVUOTDINHMUC)
        p("troCapThoiViecPhepNam") = GetCellNumber(ws, r, COL_SALARY_TROCAPTHOIVIECPHEPNAM)
        p("tongKhoanThuNhap") = GetCellNumber(ws, r, COL_SALARY_TONGKHOANTHUNHAP)

        p("bhxh") = GetCellNumber(ws, r, COL_SALARY_BHXH)
        p("bhyt") = GetCellNumber(ws, r, COL_SALARY_BHYT)
        p("bhtn") = GetCellNumber(ws, r, COL_SALARY_BHTN)
        p("thueThuNhap") = GetCellNumber(ws, r, COL_SALARY_THUETHUNHAP)
        p("tamUng") = GetCellNumber(ws, r, COL_SALARY_TAMUNG)
        p("khauTruKhac") = GetCellNumber(ws, r, COL_SALARY_KHAUTRUKHAC)
        p("luongThucLinh") = GetCellNumber(ws, r, COL_SALARY_LUONGTHUCLINH)

        Dim numberKeys As Collection
        Set numberKeys = New Collection
        Dim k As Variant
        For Each k In p.Keys
            If k <> "maNV" And k <> "hoTen" Then numberKeys.Add k
        Next k

        payroll.Add Json.DictToJson(p, numberKeys)

ContinueLoop:
    Next r
End Sub

Private Function ExcludeTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, _
                                 ByVal rawLastRow As Long) As Long
    Dim scanTo As Long
    scanTo = rawLastRow + 5

    Dim r As Long
    For r = startRow To scanTo
        Dim colCText As String
        colCText = UCase(Trim(SafeText(GetCellText(ws, r, COL_SALARY_TOTAL_MARKER))))

        If InStr(1, colCText, TOTAL_ROW_KEYWORD, vbTextCompare) > 0 Then
            LogMessage "  [Salary] Dong TOTAL tai dong " & r & "; chi doc tu dong " & _
                       startRow & " den " & (r - 1) & "."

            If Not IsYellowFill(ws.Cells(r, ColLetterToNumber(COL_SALARY_TOTAL_MARKER))) Then
                LogMessage "  [Canh bao][Salary] Dong TOTAL (dong " & r & _
                           ") khong co mau vang nhu ky vong - van loai bo theo chu TOTAL."
            End If

            ExcludeTotalRow = r - 1
            Exit Function
        End If

        If r > rawLastRow Then
            Dim manvText As String
            manvText = SafeText(GetCellText(ws, r, COL_SALARY_MANV))
            If manvText = "" And colCText = "" Then Exit For
        End If
    Next r

    ExcludeTotalRow = rawLastRow
End Function

Private Function IsYellowFill(ByVal c As Range) As Boolean
    On Error Resume Next
    Dim clr As Long
    clr = c.Interior.Color
    IsYellowFill = (clr = RGB(255, 255, 0)) Or _
                   (clr = RGB(255, 255, 153)) Or _
                   (clr = RGB(255, 192, 0))
    On Error GoTo 0
End Function

Public Function ColLetterToNumber(ByVal letters As String) As Long
    ColLetterToNumber = Range(letters & "1").Column
End Function

Private Function GetCellText(ByVal ws As Worksheet, ByVal row As Long, ByVal colLetter As String) As String
    Dim c As Range
    Set c = ws.Cells(row, ColLetterToNumber(colLetter))
    If c.NumberFormat = "@" Then
        GetCellText = CStr(c.Value)
    Else
        GetCellText = c.Text
        If GetCellText = "" Then GetCellText = CStr(c.Value)
    End If
End Function

Private Function GetCellNumber(ByVal ws As Worksheet, ByVal row As Long, ByVal colLetter As String) As Double
    Dim c As Range
    Set c = ws.Cells(row, ColLetterToNumber(colLetter))
    If IsNumeric(c.Value2) And Not IsEmpty(c.Value2) Then
        GetCellNumber = CDbl(c.Value2)
    Else
        GetCellNumber = 0
    End If
End Function

Private Function SafeText(ByVal s As String) As String
    SafeText = Trim(Replace(s, Chr(39), ""))
End Function

Private Function NormalizeName(ByVal s As String) As String
    NormalizeName = UCase(Trim(s))
End Function

' ===================== HTTP POST JSON UTF-8 =====================

Private Function PostJsonUtf8WithRetry(ByVal url As String, ByVal jsonBody As String, _
                                       ByVal label As String) As String
    Dim attempt As Long
    Dim responseText As String
    Dim code As String

    For attempt = 1 To HTTP_MAX_ATTEMPTS
        responseText = PostJsonUtf8(url, jsonBody)

        If Left$(responseText, Len("LOI GUI HTTP:")) = "LOI GUI HTTP:" Then
            code = "HTTP"
        Else
            code = ResponseCode(responseText)
            If code = "" Then code = "UNKNOWN"
        End If

        ' Thanh cong thuc su: GAS tra ok=true.
        If ResponseIsOk(responseText) Then
            PostJsonUtf8WithRetry = responseText
            Exit Function
        End If

        ' Chi retry loi tam thoi.
        If Not IsRetryableResponse(responseText, code) Then
            Err.Raise vbObjectError + 2201, "PostJsonUtf8WithRetry", _
                      "[" & label & "] GAS tu choi request: " & responseText
        End If

        If attempt < HTTP_MAX_ATTEMPTS Then
            LogMessage "[" & label & "] Lan thu " & attempt & "/" & HTTP_MAX_ATTEMPTS & _
                       " that bai (" & responseText & "). Cho " & HTTP_RETRY_WAIT_SEC & _
                       "s roi thu lai..."
            SleepSeconds HTTP_RETRY_WAIT_SEC
        End If
    Next attempt

    Err.Raise vbObjectError + 2202, "PostJsonUtf8WithRetry", _
              "[" & label & "] Khong dong bo duoc sau " & HTTP_MAX_ATTEMPTS & _
              " lan: " & responseText
End Function

Private Function PostJsonUtf8(ByVal url As String, ByVal jsonBody As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim bodyBytes() As Byte
    bodyBytes = Utf8BytesFromString(jsonBody)

    On Error GoTo HttpFailed

    http.Open "POST", url, False

    ' Google Apps Script Web App co the redirect sang URL thuc thi.
    http.Option(6) = True

    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, HTTP_TIMEOUT_RECEIVE_MS

    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "User-Agent", "PayrollSync/2.0"

    http.Send bodyBytes

    PostJsonUtf8 = http.ResponseText
    Exit Function

HttpFailed:
    PostJsonUtf8 = "LOI GUI HTTP: " & Err.Number & " - " & Err.Description
End Function

Private Function ResponseIsOk(ByVal responseText As String) As Boolean
    ResponseIsOk = (InStr(1, responseText, """ok"":true", vbTextCompare) > 0)
End Function

Private Function ResponseCode(ByVal responseText As String) As String
    Dim p As Long, q As Long, startAt As Long
    startAt = InStr(1, responseText, """code""", vbTextCompare)
    If startAt = 0 Then Exit Function

    p = InStr(startAt, responseText, ":", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + 1, responseText, """", vbBinaryCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 1, responseText, """", vbBinaryCompare)
    If q = 0 Then Exit Function

    ResponseCode = Mid$(responseText, p + 1, q - p - 1)
End Function

Private Function IsRetryableResponse(ByVal responseText As String, ByVal code As String) As Boolean
    If Left$(responseText, Len("LOI GUI HTTP:")) = "LOI GUI HTTP:" Then
        IsRetryableResponse = True
        Exit Function
    End If

    Select Case UCase$(code)
        Case "SYNC_BUSY", "SYNC_ERROR", "INTERNAL_ERROR"
            IsRetryableResponse = True
        Case Else
            IsRetryableResponse = False
    End Select
End Function

Private Sub SleepSeconds(ByVal seconds As Long)
    Dim wakeAt As Date
    wakeAt = Now + TimeSerial(0, 0, seconds)
    Application.Wait wakeAt
End Sub

Private Function Utf8BytesFromString(ByVal s As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText s
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    Utf8BytesFromString = stream.Read
    stream.Close
End Function

Public Sub LogMessage(ByVal msg As String)
    Debug.Print msg
    On Error Resume Next

    Dim logPath As String
    logPath = Environ$("TEMP") & "\SyncPayroll.log"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile(logPath, 8, True, True)
    ts.WriteLine Format$(Now, "yyyy-MM-dd HH:mm:ss") & " | " & msg
    ts.Close

    On Error GoTo 0
End Sub
