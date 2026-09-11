Attribute VB_Name = "Json"
'-------------------------------------------------------------------------
' Json.bas
' Bo ham JSON toi gian, tu viet, khong phu thuoc thu vien ngoai (khong can
' cai VBA-JSON / JsonConverter). Du dung cho cau truc du lieu co dinh cua
' project nay: 1 object goc chua "employees" va "payroll" la 2 mang object.
'
' Cach dung:
'   Dim emp As Object: Set emp = CreateObject("Scripting.Dictionary")
'   emp("maNV") = "NV001"
'   emp("hoTen") = "Nguyen Van A"
'   ...
'   JsonEscape / JsonNumber / DictToJson / CollectionToJsonArray ben duoi.
'-------------------------------------------------------------------------
Option Explicit

' Escape 1 chuoi de dua vao JSON string an toan (dau ngoac kep, backslash,
' xuong dong, tab, ky tu Unicode tieng Viet giu nguyen vi JSON ho tro UTF-8/
' \uXXXX, va WinHTTP se gui body dang UTF-8).
Public Function JsonEscape(ByVal s As String) As String
    Dim result As String
    result = s
    result = Replace(result, "\", "\\")
    result = Replace(result, Chr(34), "\" & Chr(34))
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JsonEscape = result
End Function

' Tra ve chuoi JSON cho 1 gia tri string: "..."
Public Function JsonString(ByVal s As String) As String
    JsonString = Chr(34) & JsonEscape(s) & Chr(34)
End Function

' Tra ve chuoi JSON cho 1 gia tri so (dung dau cham thap phan bat ke Regional
' Settings cua may, tranh loi khi may dung dau phay lam dau thap phan).
Public Function JsonNumber(ByVal v As Double) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ",", ".")
    JsonNumber = s
End Function

' Chuyen 1 Scripting.Dictionary (key => value, value la String hoac Double)
' thanh 1 JSON object. valueIsNumberKeys: mang cac key can xuat o dang so
' (khong bao trong dau ngoac kep).
Public Function DictToJson(ByVal dict As Object, ByVal numberKeys As Collection) As String
    Dim parts() As String
    Dim i As Long
    Dim key As Variant
    Dim n As Long
    n = dict.Count
    ReDim parts(0 To n - 1)
    i = 0
    For Each key In dict.Keys
        Dim isNumber As Boolean
        isNumber = CollectionContains(numberKeys, CStr(key))
        If isNumber Then
            parts(i) = JsonString(CStr(key)) & ":" & JsonNumber(CDbl(dict(key)))
        Else
            parts(i) = JsonString(CStr(key)) & ":" & JsonString(CStr(dict(key)))
        End If
        i = i + 1
    Next key
    DictToJson = "{" & Join(parts, ",") & "}"
End Function

' Noi 1 Collection cac chuoi JSON object (da duoc dung sanh boi DictToJson)
' thanh 1 JSON array: [ {...}, {...} ]
Public Function CollectionToJsonArray(ByVal items As Collection) As String
    Dim parts() As String
    Dim n As Long
    n = items.Count
    If n = 0 Then
        CollectionToJsonArray = "[]"
        Exit Function
    End If
    ReDim parts(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        parts(i - 1) = items(i)
    Next i
    CollectionToJsonArray = "[" & Join(parts, ",") & "]"
End Function

Private Function CollectionContains(ByVal col As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In col
        If CStr(item) = value Then
            CollectionContains = True
            Exit Function
        End If
    Next item
    CollectionContains = False
End Function
