Sub 链接22222222222222()
'选中一个单元格，以此为起点生成目录表，并设置好链接
    Dim rng As Range
    Dim everyRng As Range
    Set rng = Selection
    '不准选择多行
    If rng.Count > 1 Then Exit Sub
    
    Dim sht As Worksheet
    Dim arr() As Variant
    '激活工作表，不用被加入目录表，arr用来存放所有工作表名称
    n = Worksheets.Count
    If n = 1 Then Exit Sub
    ReDim arr(1 To n - 1)
    i = 1
    For Each sht In Worksheets
        If sht.Name <> ActiveSheet.Name Then
            arr(i) = sht.Name
            i = i + 1
        End If
    Next
    '输出目录
    rng.Resize(n - 1, 1) = Application.WorksheetFunction.Transpose(arr)
    '生成链接
    For Each everyRng In rng.Resize(n - 1, 1)
        everyRng.Select
        Selection.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & everyRng.Value & "'!A1", ScreenTip:="", TextToDisplay:=everyRng.Value
    Next

End Sub
