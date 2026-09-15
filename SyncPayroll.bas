Attribute VB_Name = "SyncPayroll"
'-------------------------------------------------------------------------
' SyncPayroll.bas
' Macro chinh: xac dinh thang luong can lay (thang truoc thang hien tai),
' tim file .xlsb tuong ung cua 2 xuong Snack va Flexible, doc du liu tu
' sheet Salary + DSCNV, ghep noi, dong goi JSON va day len Google Apps
' Script (endpoint /sync).
'
' Cach chay tu dong hang ngay 10 (hoac hang ngay, script se tu bo qua neu
' thang do da duoc dong bo va khong co thay doi can thiet - xem ghi chu o
' cuoi ham RunSync): xem docs/WINDOWS_TASK_SCHEDULER.md
'-------------------------------------------------------------------------
Option Explicit

Public Sub RunSync()
    Dim targetMonth As Integer, targetYear As Integer
    GetTargetMonthYear targetMonth, targetYear

    Dim thangStr As String
    thangStr = Format(targetMonth, "00") & "-" & CStr(targetYear)

    LogMessage "=== Bat dau dong bo luong thang " & thangStr & " luc " & Now & " ==="

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    SyncOneWorkshop SNACK_ROOT, SNACK_FILE_PREFIX, SNACK_LABEL, targetMonth, targetYear, thangStr
    SyncOneWorkshop FLEXIBLE_ROOT, FLEXIBLE_FILE_PREFIX, FLEXIBLE_LABEL, targetMonth, targetYear, thangStr

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    LogMessage "=== Ket thuc dong bo luc " & Now & " ==="
End Sub

' Thang can lay = thang truoc thang hien tai (vd hom nay la 9/2026 -> lay 8/2026)
Private Sub GetTargetMonthYear(ByRef outMonth As Integer, ByRef outYear As Integer)
    Dim today As Date
    today = Date
    Dim prevMonthDate As Date
    prevMonthDate = DateAdd("m", -1, today)
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
        LogMessage "[" & label & "] KHONG TIM THAY FILE. Bo qua xuong nay."
        Exit Sub
    End If

    Dim wb As Workbook
    On Error GoTo OpenFailed
    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=True, _
                             Password:=XLSB_PASSWORD, IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0

    Dim wsSalary As Worksheet, wsDscnv As Worksheet
    Set wsSalary = GetSheetOrNothing(wb, SHEET_SALARY)
    Set wsDscnv = GetSheetOrNothing(wb, SHEET_DSCNV)

    If wsSalary Is Nothing Or wsDscnv Is Nothing Then
        LogMessage "[" & label & "] Khong tim thay sheet Salary hoac DSCNV trong file."
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    Dim dscnvLookup As Object
    Set dscnvLookup = BuildDscnvLookup(wsDscnv)

    Dim employees As Collection, payroll As Collection
    Set employees = New Collection
    Set payroll = New Collection
    ReadSalarySheet wsSalary, dscnvLookup, employees, payroll

    wb.Close SaveChanges:=False

    LogMessage "[" & label & "] Doc duoc " & payroll.Count & " dong luong, " & employees.Count & " nhan vien."

    If payroll.Count = 0 Then
        LogMessage "[" & label & "] Khong co du lieu de dong bo."
        Exit Sub
    End If

    Dim jsonBody As String
    jsonBody = "{" & _
        Json.JsonString("action") & ":" & Json.JsonString("sync") & "," & _
        Json.JsonString("apiKey") & ":" & Json.JsonString(SYNC_API_KEY) & "," & _
        Json.JsonString("xuong") & ":" & Json.JsonString(label) & "," & _
        Json.JsonString("thang") & ":" & Json.JsonString(thangStr) & "," & _
        Json.JsonString("employees") & ":" & Json.CollectionToJsonArray(employees) & "," & _
        Json.JsonString("payroll") & ":" & Json.CollectionToJsonArray(payroll) & _
        "}"

    Dim responseText As String
    responseText = PostJsonUtf8WithRetry(GAS_URL, jsonBody, label)
    LogMessage "[" & label & "] Phan hoi tu server: " & responseText

    Exit Sub

OpenFailed:
    LogMessage "[" & label & "] LOI khi mo file (kiem tra mat khau/duong dan): " & Err.Description
    On Error GoTo 0
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
            If Not result.Exists(key) Then
                result.Add key, info
            End If
        End If
    Next r

    Set BuildDscnvLookup = result
End Function

' Doc sheet Salary tu dong DATA_START_ROW, ghep voi DSCNV lookup theo ten,
' ghi vao 2 Collection: employees (String JSON, khong trung MaNV) va
' payroll (String JSON, 1 dong / nhan vien).
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
            LogMessage "  [Canh bao] Khong khop DSCNV cho nhan vien: " & hoTen & " (Ma NV " & maNV & ")"
        End If

        ' ---- employees (chi them 1 lan cho moi Ma NV) ----
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

        ' ---- payroll ----
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

' ===================== Xac dinh & loai dong TOTAL =====================

' Do tu startRow den rawLastRow (co du phong them vai dong phia duoi de bat
' truong hop dong TOTAL co gia tri o cot Ma NV khien End(xlUp) tinh du vao no),
' tim dong dau tien co chu "TOTAL" trong cot COL_SALARY_TOTAL_MARKER (cot C).
' Neu tim thay, chi doc du lieu den ngay TRUOC dong do. Neu khong tim thay
' (vd file khong co dong TOTAL), giu nguyen rawLastRow nhu truoc.
Private Function ExcludeTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, _
                                  ByVal rawLastRow As Long) As Long
    Dim scanTo As Long
    scanTo = rawLastRow + 5 ' du phong vai dong phia sau phong khi TOTAL nam ngoai rawLastRow

    Dim r As Long
    For r = startRow To scanTo
        Dim colCText As String
        colCText = UCase(Trim(SafeText(GetCellText(ws, r, COL_SALARY_TOTAL_MARKER))))

        If InStr(1, colCText, TOTAL_ROW_KEYWORD, vbTextCompare) > 0 Then
            LogMessage "  [Salary] Dong TOTAL tai dong " & r & "; chi doc tu dong " & _
                       startRow & " den " & (r - 1) & "."

            If Not IsYellowFill(ws.Cells(r, ColLetterToNumber(COL_SALARY_TOTAL_MARKER))) Then
                LogMessage "  [Canh bao][Salary] Dong TOTAL (dong " & r & _
                           ") khong co mau vang nhu ky vong - van loai bo dong nay theo chu TOTAL."
            End If

            ExcludeTotalRow = r - 1
            Exit Function
        End If

        ' Neu qua rawLastRow ma van gap dong trong lien tiep (khong phai TOTAL)
        ' thi dung do som de tranh quet vo tan.
        If r > rawLastRow Then
            Dim manvText As String
            manvText = SafeText(GetCellText(ws, r, COL_SALARY_MANV))
            If manvText = "" And colCText = "" Then Exit For
        End If
    Next r

    ' Khong tim thay dong TOTAL -> giu nguyen ranh gioi cu (khong loai gi them).
    ExcludeTotalRow = rawLastRow
End Function

' Dong TOTAL thuong duoc to mau vang trong file nguon. Chi dung de CANH BAO khi
' khong khop (vd sheet doi mau format), khong dung de quyet dinh loai dong.
Private Function IsYellowFill(ByVal c As Range) As Boolean
    On Error Resume Next
    Dim clr As Long
    clr = c.Interior.Color

    ' Cac tone vang thuong gap: vang chuan, vang nhat, vang/cam dam (accent)
    IsYellowFill = (clr = RGB(255, 255, 0)) Or _
                   (clr = RGB(255, 255, 153)) Or _
                   (clr = RGB(255, 192, 0))
    On Error GoTo 0
End Function

' ===================== Helpers doc o (cell) =====================

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
    SafeText = Trim(Replace(s, Chr(39), "")) ' bo dau nhay don ' o dau neu co
End Function

Private Function NormalizeName(ByVal s As String) As String
    NormalizeName = UCase(Trim(s))
End Function

' ===================== Goi HTTP (POST JSON, UTF-8) =====================

' Goi PostJsonUtf8 toi da HTTP_MAX_ATTEMPTS lan. Thu lai khi: loi ket noi/
' timeout ("LOI GUI HTTP:"), hoac server tra ve SYNC_BUSY (dang co phien dong
' bo khac - hiem khi xay ra vi GAS da ghi theo lo rat nhanh, nhung 2 xuong co
' the goi len gan nhu cung luc), hoac loi he thong tam thoi (SYNC_ERROR /
' INTERNAL_ERROR). Cac loi khac (vd sai API key - FORBIDDEN) KHONG thu lai vi
' thu lai cung se that bai.
Private Function PostJsonUtf8WithRetry(ByVal url As String, ByVal jsonBody As String, _
                                        ByVal label As String) As String
    Dim attempt As Long
    Dim responseText As String

    For attempt = 1 To HTTP_MAX_ATTEMPTS
        responseText = PostJsonUtf8(url, jsonBody)

        If Left(responseText, Len("LOI GUI HTTP:")) <> "LOI GUI HTTP:" And _
           InStr(1, responseText, """code"":""SYNC_BUSY""", vbTextCompare) = 0 And _
           InStr(1, responseText, """code"":""SYNC_ERROR""", vbTextCompare) = 0 And _
           InStr(1, responseText, """code"":""INTERNAL_ERROR""", vbTextCompare) = 0 Then
            PostJsonUtf8WithRetry = responseText
            Exit Function
        End If

        If attempt < HTTP_MAX_ATTEMPTS Then
            LogMessage "[" & label & "] Lan thu " & attempt & "/" & HTTP_MAX_ATTEMPTS & _
                       " that bai (" & responseText & "). Cho " & HTTP_RETRY_WAIT_SEC & _
                       "s roi thu lai..."
            SleepSeconds HTTP_RETRY_WAIT_SEC
        End If
    Next attempt

    PostJsonUtf8WithRetry = responseText & " (da thu lai " & HTTP_MAX_ATTEMPTS & " lan)"
End Function

Private Function PostJsonUtf8(ByVal url As String, ByVal jsonBody As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim bodyBytes() As Byte
    bodyBytes = Utf8BytesFromString(jsonBody)

    On Error GoTo HttpFailed
    http.Open "POST", url, False
    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, HTTP_TIMEOUT_CONNECT_MS, _
                      HTTP_TIMEOUT_SEND_MS, HTTP_TIMEOUT_RECEIVE_MS
    http.SetRequestHeader "Content-Type", "text/plain;charset=utf-8"
    http.Send bodyBytes
    PostJsonUtf8 = http.responseText
    Exit Function

HttpFailed:
    PostJsonUtf8 = "LOI GUI HTTP: " & Err.Description
    On Error GoTo 0
End Function

' Application.Wait can gia tri Date tuyet doi nen phai cong don tu Now(); voi
' so giay lon (>~24h) se tran, nhung o day chi dung vai chuc giay nen an toan.
Private Sub SleepSeconds(ByVal seconds As Long)
    Dim wakeAt As Date
    wakeAt = Now + TimeSerial(0, 0, seconds)
    Application.Wait wakeAt
End Sub

' Chuyen 1 chuoi VBA (UTF-16) thanh mang byte UTF-8 (bo BOM), dung ADODB.Stream.
Private Function Utf8BytesFromString(ByVal s As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText s
    stream.Position = 0
    stream.Type = 1 ' adTypeBinary
    stream.Position = 3 ' bo qua 3 byte BOM cua UTF-8
    Utf8BytesFromString = stream.Read
    stream.Close
End Function

' ===================== Ghi log =====================

Public Sub LogMessage(ByVal msg As String)
    Debug.Print msg
    On Error Resume Next
    Dim logPath As String
    logPath = Environ("TEMP") & "\SyncPayroll.log"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    ' Format:=True (TristateTrue) => ghi file dang Unicode (UTF-16), hien thi
    ' dung tieng Viet co dau bat ke may dang dung ma trang (code page) nao.
    ' Truoc day dung Print # se ghi theo ANSI he thong, neu may khong dat
    ' locale Tieng Viet thi cac ky tu co dau se bi thay bang dau "?".
    Set ts = fso.OpenTextFile(logPath, 8, True, True)
    ts.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & " | " & msg
    ts.Close
    On Error GoTo 0
End Sub
