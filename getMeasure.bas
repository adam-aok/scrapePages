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


'similar syntax is used for a Power Query M formula. See here
'if Text.Contains([SquareFootage.1],"sf") then [SquareFootage.1]
'else
'if Text.Contains([SquareFootage.2],"sf") then [SquareFootage.2]
'else
'if Text.Contains([SquareFootage.3],"sf") then [SquareFootage.3]
'else
'if Text.Contains([SquareFootage.4],"sf") then [SquareFootage.4]
'else
'if Text.Contains([SquareFootage.5],"sf") then [SquareFootage.5]
'else
'if Text.Contains([SquareFootage.6],"sf") then [SquareFootage.6]
'else
'if Text.Contains([SquareFootage.7],"sf") then [SquareFootage.7]
'else
'if Text.Contains([SquareFootage.8],"sf") then [SquareFootage.8]
'else
'if Text.Contains([SquareFootage.9],"sf") then [SquareFootage.9]
'else
'0
