Sub Circulation()

Dim row As Long
Dim num As Long
Dim init_num As Long
Dim MAX As Long
Dim Column As String
Dim Circ As Long
Dim head As Long


'ここから↓設定項目(右辺を置き換える)
init_num = 1 '最小入力数字
Min = 1 '何行目から
MAX = 100 '何行目まで繰り返す
Column = "A" 'どこ列で実行するか("は消さんといて)
Circ = 5 '何回ずつ数字を繰り返すか
head = 10 'いくつまで繰り返すか
'設定項目ここまで


num = init_num
head = head + 1

    For row = 1 To MAX
        Cells(row, Column).Value = num

        If row Mod Circ = 0 Then
           num = num + 1

        End If
    
    If num Mod head = 0 Then
        num = init_num
    End If
    
    Next row
       
End Sub
