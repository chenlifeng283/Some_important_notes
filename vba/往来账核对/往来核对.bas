
Sub jcgsmc() '检查公司名称输入
Dim i As Integer
i = 3
Do While Cells(i, 2) <> "" '遍历B列,有空格就停止

        If Cells(i, 2) <> Cells(2, i) Then  '表格中对比的单位名称的规律是行列相反，因此只需一个变量
        Cells(i, 2).Interior.Color = RGB(127, 255, 170) '单位名称不相同，就着色
        Cells(2, i).Interior.Color = RGB(127, 255, 170)
        
        End If
        i = i + 1 '循环到下一个数
Loop
'下面这段if判断比较怪，颜色的判断，如果直接用if……then居然判断不出来，一定要加else,才能达到效果，为什么？
 If Columns(2).Interior.ColorIndex = xlNone Then '如是B列没有着色，就不做任何操作
 Else
     MsgBox "行列单位不匹配，已标记颜色" '如果B列有着色，就用提示框提示
 End If

End Sub


Sub check() ' 金额核对
Dim i, j As Integer
i = 3
Do While Cells(i, 2) <> "" '遍历第二列，有单位就判断它的取值
    j = 3
    Do While Cells(2, j) <> "" '遍历第二行，程序顺序是:第3行，第3……j列，第4行，第3……J列，依次……
        If Cells(i, j).Value <> -Cells(j, i).Value Then '表格中的规律是两家单位分别作为对方的二级科目，在单元格中行列是相反的，如3行4列和4行3列的数值就是希望去对比的数值。
        Cells(i, j).Interior.Color = RGB(218, 150, 148) '不相同就着色
        Cells(j, i).Interior.Color = RGB(218, 150, 148)
        End If
    j = j + 1
    Loop
i = i + 1
Loop
'以下，数字区域是否有着色的判断
If Range(Cells(3, 3), Cells(100, 100)).Interior.ColorIndex = xlNone Then
MsgBox "数据全部准确"
Else
MsgBox "数据有误，已标记颜色"
End If
            
End Sub

Sub clear() '改正后重新判断，需先清理数据（把标色去掉）
 Application.ScreenUpdating = False


With Range("b2").CurrentRegion
    '.ClearContents '不能清除内容，有公式，只需清除相应有错误的着色，便于重新计算后发现新情况。
    .NumberFormat = "0.00"
    .Interior.ColorIndex = xlNone
    .Font.Underline = xlNone
    .Font.Color = RGB(0, 0, 0)
End With
 Application.ScreenUpdating = True
End Sub
