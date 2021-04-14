VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Splitting rows by semicolons for Great Exhibition scraped data--by Adam Keefe, 05/26/2020
Sub toSplit()
Dim v As Integer, h As Integer, rowLast As Integer, totalRowsFirst As Integer, totalRowsSecond As Integer, secondCount As Integer
Dim pn As String, serv As String, soft As String
Dim nameArr As Variant
'rowArray() As Variant

nameArr = Array()

'get total number of rows
totalRowsFirst = Sheet3.Rows(Rows.Count).End(xlUp).Row

'run once for splitting
For v = 2 To totalRowsFirst
    soft = Sheet3.Cells(v, 5).Value
    If InStr(soft, ";") <> 0 Then
        nameArr = Split(soft, ";")
        For h = 0 To UBound(nameArr)
            rowLast = Sheet4.Rows(Rows.Count).End(xlUp).Row + 1
            Sheet3.Rows(v).Copy
            Sheet4.Rows(rowLast).Insert
         '   Sheet3.Rows(rowLast) = Sheet2.Rows(v)
            Sheet4.Cells(rowLast, 5) = nameArr(h)
        Next h
    End If
    If InStr(soft, ";") = 0 Then
        rowLast = Sheet4.Rows(Rows.Count).End(xlUp).Row + 1
        Sheet3.Rows(v).Copy
        Sheet4.Rows(rowLast).Insert
    End If
    
Next v

End Sub
