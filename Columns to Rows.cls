VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub toRows()
Dim v As Integer, h As Integer, rowLast As Integer, totalRowsFirst As Integer, totalRowsSecond As Integer
Dim pn As String, serv As String
'rowArray() As Variant
totalRowsFirst = Sheet1.Rows(Rows.Count).End(xlUp).Row

For v = 2 To totalRowsFirst
    pn = Sheet1.Cells(v, 1)
    For h = 2 To 6
        If Sheet1.Cells(v, h) <> "" Then
        serv = Sheet1.Cells(v, h)
        rowLast = Sheet2.Rows(Rows.Count).End(xlUp).Row + 1
        Sheet2.Cells(rowLast, 1) = pn
        Sheet2.Cells(rowLast, 1) = serv
        End If
    Next h
Next v

End Sub
