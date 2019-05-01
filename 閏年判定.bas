Attribute VB_Name = "Module1"
Option Explicit
Sub 閏年判定()
    Dim year
    'Cells()
    Cells(1, 2).Clear
    
    'Cells(1, 1)を変数yearに格納
    year = Cells(1, 1)
    
    '結果を計算、表示
    If year Mod 900 = 200 Or year Mod 900 = 600 Then
        Cells(1, 2) = "閏年"
    ElseIf year Mod 100 = 0 Then
        Cells(1, 2) = "平年"
    ElseIf year Mod 4 = 0 Then
        Cells(1, 2) = "閏年"
    Else
        Cells(1, 2) = "平年"
        
    End If

End Sub

