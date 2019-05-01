Attribute VB_Name = "Module1"
Option Explicit

Sub 最大値判定()
 
  Dim Maxval As Long, minVal As Long
  
  Cells(2, 1).Clear
 
  'Excelのワークシート関数を使う
  With Application.WorksheetFunction
 
  'セルB2～H8までのセルに記入された数値の最大値, 最小値を求める
       Maxval = .MAX(Range(Cells(1, 1), Cells(1, 10)))
       Cells(2, 1) = Maxval
  End With
  
End Sub
