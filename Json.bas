Attribute VB_Name = "Json"
Option Explicit

Public Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, Chr(34), "\" & Chr(34))
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonEscape = t
End Function

Public Function JsonString(ByVal s As String) As String
    JsonString = Chr(34) & JsonEscape(s) & Chr(34)
End Function

Public Function JsonNumber(ByVal v As Double) As String
    Dim s As String
    If IsNaNNumber(v) Then
        JsonNumber = "0"
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, ",", ".")
    If s = "" Or s = "-" Then s = "0"
    JsonNumber = s
End Function

Private Function IsNaNNumber(ByVal v As Double) As Boolean
    IsNaNNumber = False
End Function

Public Function DictToJson(ByVal dict As Object, ByVal numberKeys As Collection) As String
    Dim parts() As String, i As Long, key As Variant, n As Long
    n = dict.Count
    If n = 0 Then DictToJson = "{}": Exit Function
    ReDim parts(0 To n - 1)
    i = 0
    For Each key In dict.Keys
        If CollectionContains(numberKeys, CStr(key)) Then
            parts(i) = JsonString(CStr(key)) & ":" & JsonNumber(CDbl(dict(key)))
        Else
            parts(i) = JsonString(CStr(key)) & ":" & JsonString(CStr(dict(key)))
        End If
        i = i + 1
    Next key
    DictToJson = "{" & Join(parts, ",") & "}"
End Function

Public Function CollectionToJsonArray(ByVal items As Collection) As String
    Dim parts() As String, i As Long
    If items.Count = 0 Then CollectionToJsonArray = "[]": Exit Function
    ReDim parts(0 To items.Count - 1)
    For i = 1 To items.Count: parts(i - 1) = CStr(items(i)): Next i
    CollectionToJsonArray = "[" & Join(parts, ",") & "]"
End Function

Private Function CollectionContains(ByVal col As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In col
        If CStr(item) = value Then CollectionContains = True: Exit Function
    Next item
    CollectionContains = False
End Function
