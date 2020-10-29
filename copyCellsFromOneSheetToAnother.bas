REM  *****  BASIC  *****

Private Sub Copy_cells_from_one_sheet_to_another()
Dim Cell_1 As String, Cell_2 As Integer, Cell_3 As Double, Cell_4 As Integer
Worksheets("Sheet1").Select
Cell_1 = Range("B2")
Cell_2 = Range("B3")
Cell_3 = Range("B4")
Cell_4 = Range("B5")
Worksheets("Sheet2").Select
Worksheets("Sheet2").Range("A1").Select
If Worksheets("Sheet2").Range("A1").Offset(1, 0) <> "" Then
Worksheets("Sheet2").Range("A1").End(xlDown).Select
End If
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Cell_1
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Cell_2
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Cell_3
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Cell_4
Worksheets("Sheet1").Select
Worksheets("Sheet1").Range("B2").Select
End Sub

