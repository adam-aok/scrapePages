Function GetMeasure(parseCol As String, argUnit As String) As String
Dim Result As String
Dim i As Integer
Dim colSplit() As String
'colSplit() = Split(parseCol, ";")
If InStr(1, parseCol, ";") <> 0 Then
colSplit() = Split(parseCol, ";")
Else
colSplit() = Split(parseCol, " (")
End If
'Result = Split(parseCol, ";")(1)
For i = 0 To UBound(colSplit)
If InStr(1, LCase(colSplit(i)), argUnit) <> 0 Then
Result = Trim(colSplit(i))
Result = Replace(Replace(Trim(colSplit(i)), "(", ""), ")", "")
End If
Next i
'Result = Replace(Replace(Trim(colSplit(i)), "(", ""), ")", "")
GetMeasure = Result
End Function
