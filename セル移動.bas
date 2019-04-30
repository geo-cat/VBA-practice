Option Explicit
Sub セル移動()
  Dim a1
  a1 = Cells(1, 3)
  Cells(1, 3).Cut
  Cells(1, 2).Cut Cells(1, 3)
  Cells(1, 1).Cut Cells(1, 2)
  Cells(1, 1) = a1
End Sub
