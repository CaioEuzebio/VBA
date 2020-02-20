Sub RemoveDuplicateRows()

Dim MyRange As Range
Dim LastRow As Long

LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Set MyRange = ActiveSheet.Range("A1:D" & LastRow)
MyRange.RemoveDuplicates Columns:=3, Header:=xlYes
End Sub
