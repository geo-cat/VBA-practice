Attribute VB_Name = "Module1"
Option Explicit
Sub ƒZƒ‹ˆÚ“®()
  Dim a1
  a1 = Cells(1, 3)
  Cells(1, 3).Cut
  Cells(1, 2).Cut Cells(1, 3)
  Cells(1, 1).Cut Cells(1, 2)
  Cells(1, 1) = a1
End Sub
