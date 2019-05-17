Sub 網掛け()

'全てのセルの選択
  With Cells
    ' 網掛けの種類を選択（パターンはたくさんある）
    .Interior.Pattern = xlGray25
       
    ' 網掛けの色のクリア実験した後にセルをまっさらにするのに使える
    '.Interior.Pattern = xlNone
  End With

End Sub
