Attribute VB_Name = "Module1"
Option Explicit

Sub �ő�l����()
 
  Dim Maxval As Long, minVal As Long
  
  Cells(2, 1).Clear
 
  'Excel�̃��[�N�V�[�g�֐����g��
  With Application.WorksheetFunction
 
  '�Z��B2�`H8�܂ł̃Z���ɋL�����ꂽ���l�̍ő�l, �ŏ��l�����߂�
       Maxval = .MAX(Range(Cells(1, 1), Cells(1, 10)))
       Cells(2, 1) = Maxval
  End With
  
End Sub
