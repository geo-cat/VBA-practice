Attribute VB_Name = "Module1"
Option Explicit

Sub Sample1()
    '変数”ans”と”num”の宣言
    Dim ans
    Dim num
    
    'ループ
    For num = 1 To 10
        '変数ansにInputBoxに入力されたものを格納する。InputBoxにはCells(num, 1)を表示する。
        ans = InputBox(Cells(num, 1))
        'if分でCells(num, 2)と比較して、一致すれば正解そうでなければ間違いと出力する。
        If ans = Cells(num, 2) Then
            MsgBox "正解"
        Else
            MsgBox "間違い"
        End If
    Next num
End Sub
