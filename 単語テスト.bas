Attribute VB_Name = "Module1"
Option Explicit
Sub Sample1()
    Dim ans
    Dim num
    
    For num = 1 To 10
        ans = InputBox(Cells(num, 1))
        
        If ans = Cells(num, 2) Then
            MsgBox "ê≥â"
        Else
            MsgBox "ä‘à·Ç¢"
        End If
    Next num
End Sub
