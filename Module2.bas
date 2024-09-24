Attribute VB_Name = "Module2"
Option Explicit
Sub Fill()
Dim nc As Integer
Dim i As Integer
Dim a As Variant
Worksheets("Sheet5").Activate
nc = WorksheetFunction.CountA(Range("2:2"))
For i = 4 To nc
a = Cells(1, (i - 1)).Value
If Cells(1, i).Value = "" Then Cells(1, i).Value = a
Range("1:1").NumberFormat = "dd/mm/yyyy"
Next i


End Sub
