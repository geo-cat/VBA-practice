Attribute VB_Name = "Module1"
Option Explicit

Sub [N»θ()
    Dim year As Integer, str As String
    'Cells()
    Cells(1, 2).Clear
    
    'Cells(1, 1)πΟyearΙi[
    year = Cells(1, 1)
    
    'ΚπvZA\¦
    If year Mod 900 = 200 Then
        Cells(1, 2) = "[N"
    ElseIf year Mod 100 = 0 Then
        Cells(1, 2) = "½N"
    ElseIf year Mod 4 = 0 Then
        Cells(1, 2) = "[N"
    Else
        Cells(1, 2) = "½N"
        
    End If

End Sub
