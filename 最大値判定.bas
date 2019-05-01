Attribute VB_Name = "Module1"
Option Explicit


Sub 最大値判定()
 '変数Maxvalの宣言
 Dim Maxval As Long, minVal As Long
  
 Cells(2, 1).Clear
 
  'Excelのワークシート関数を使用する
  With Application.WorksheetFunction
 
  'Cells(1, 1)からCells(1, 10)までの数値の最大値を求め、変数Maxvalに格納する。
       Maxval = .MAX(Range(Cells(1, 1), Cells(1, 10)))
 'Cells(2, 1)にMaxvalを出力
       Cells(2, 1) = Maxval
  End With
  
End Sub
