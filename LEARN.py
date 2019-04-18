2019.04.08 完成注会报名
2019.04.08 经济法第七章证券法  第八章破产法 中华题库基础版第四章
2019.04.09 经济法第八章破产法  第九章票据法
2019.04.11 第八、九章看完
2019.04.12 第十一章垄断法
2019.04.13 第十一章垄断法看完
2019.04.14 第十二章涉外经济法12/24
2019.04.16 第十二章涉外经济法 看完
2019.04.17 第十二章国有企业法11/22



Sub testUnion()
lRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim rng As Range
Dim rng1 As Range
Dim rng2 As Range

Dim a(0 To 99999) As Integer
Dim b As Long
b = 0
a(1) = 0

Set rng1 = Cells(lRow, 1)
For i = 1 To lRow Step 1
    If Cells(i, 1) = "奥迪" Then
        a(b) = i
        b = b + 1
    End If
Next

For bb = 0 To b Step 1
    If a(bb) <> 0 Then
        If bb = 0 Then
            Set rng1 = Cells(a(bb), 1)
            Set rng = Union(rng1, rng1)
        Else
            Set rng1 = Cells(a(bb), 1)
            Set rng = Union(rng, rng1)
        End If
    
        
    End If
Next

rng.Select

End Sub

