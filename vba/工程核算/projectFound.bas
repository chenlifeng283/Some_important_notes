知识点：1.range("A35670").end(xlup).offset(1,0)定位有数据的最后一个单元格，并下移一格，开始写数据
	   2.设置“超链接” selection.Hyperlinks.add 
	   3.inputbox 输入对话框
	   4.instr("ABC","A") 字符串ABC中是否包含"A",返回的是”个数“，因为0是false,其中整数是True,所以可以用来判断。
	   5.split("A,B,C",",") 以逗号为分隔符拆分字符串，返回的是一个数组，arrary("A","B","C"),如arrary(0) = "A"
	   6.窗体，textbox1.value是用户输入的数据。
	   

Option Explicit
Public sht1 As Worksheet, sht As Worksheet, o, p, q, r

Sub 更新目录表()
Dim rng As Range, q
    Set sht1 = Worksheets("目录")
    sht1.Range("B3:K5000").ClearContents
    '为了随时更新且不会重得，每次都是删除全数据，重新写，有个缺点，‘备注’栏不是取自各’明细表‘，如果填写了备注，
    '后面又新插入的表，备注栏就是错位的，不过用表中’增加明细表‘按键就不会出错
    
    For Each sht In Worksheets
        If sht.Name <> sht1.Name And sht.Name <> "名称管理器" Then '遍历除了‘目录’、‘名称管理器’的所有表
            Set rng = sht1.Range("B5000").End(xlUp).Offset(1, 0)  ' 定位Ｂ列有数据的最后一个单元格，并下移一格，开始写数据
            rng = Mid(sht.Range("A2"), 6)  '从各工程明细表取工程名，填入目录表
            q = rng.Value
            sht1.Activate
            rng.Select
            
            On Error Resume Next:
            Selection.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sht.Name & "!A1", ScreenTip:="", TextToDisplay:=rng.Value
            With rng   '依次填入数据
                .Offset(0, 1) = Mid(sht.Range("H2"), 6)
                .Offset(0, 2) = Mid(sht.Range("K2"), 7)
                .Offset(0, 3) = sht.Range("N4")
                .Offset(0, 4) = sht.Range("O4")
                .Offset(0, 5) = sht.Range("K4")
                .Offset(0, 6) = sht.Range("J4")
                .Offset(0, 7) = sht.Range("M4")
                .Offset(0, 8) = sht.Range("Q4")
                .Offset(0, 9) = sht.Range("P4")
            End With
                
        End If
    Next sht
    
    
    MsgBox "更新完成!!", , "工程目录更新"

End Sub


Sub add_worksheets11() ' 按要求创建新表
Dim sht3 As Worksheet
Dim c As Variant
Dim arr1
    Worksheets(Sheet3).Copy after:=Worksheets(Worksheets.Count)  '复制sheet3,在最后位置
    Set sht3 = ActiveSheet  ' 新建表是被马上激活的，赋值给sht3
    sht3.Range("A2") = Left(sht3.Range("A2"), 5) '从旧表复制过来，表头等有数据，按要求保留、删除
    sht3.Range("H2") = Left(sht3.Range("H2"), 5)
    sht3.Range("K2") = Left(sht3.Range("K2"), 6)
    sht3.Range("P2") = Left(sht3.Range("P2"), 5)
    sht3.Range("A5:P5000").ClearContents  ' 表中数据全删
    
    On Error GoTo x:
    ' 当输入的工作表名称重复，或者不输名称关掉对话框都会跳错误，
    ' 这时就转到退出sub,新建的工作表仍有效，表名称就是系统自动生成的表名
    
    c = InputBox(prompt:="依次输入工作表名、工程名称、合同金额、承包人、工程时间，用逗号隔开！", Title:="创建工作表") '跳出对话框重命名表名
    If InStr(c, ",") Then
        arr1 = Split(c, ",")
        sht3.Name = arr1(0)
        sht3.Range("A2") = Left(sht3.Range("A2"), 5) & arr1(1)
        sht3.Range("H2") = Left(sht3.Range("H2"), 5) & arr1(2)
        sht3.Range("K2") = Left(sht3.Range("K2"), 6) & arr1(3)
        sht3.Range("P2") = Left(sht3.Range("P2"), 5) & arr1(4)
        
        
    Else:
        sht3.Name = c
    End If
    
x:    Exit Sub

'On Error GoTo x:
'
'    c = InputBox(prompt:="输入工作表名！", Title:="创建工作表")
'    If c = "" Then
'        sht3.Name = Application.WorksheetFunction.RandBetween(1, 100) & "连名称都不想改的懒人" & Application.WorksheetFunction.RandBetween(1, 100)
'
'    Else:
'        sht3.Name = c
'    End If
'
'    Exit Sub
'
'x:
'    MsgBox "工作表名不能重复"
'    Exit Sub
    
End Sub



Sub add_worksheets22()   '有窗体输入工作表名等
Dim sht3 As Worksheet
Dim c As Variant
    Worksheets(Sheet3).Copy after:=Worksheets(Worksheets.Count)  '复制sheet3,在最后位置
    Set sht3 = ActiveSheet  ' 新建表是被马上激活的，赋值给sht3
    
    UserForm1.Show vbModal
    
    sht3.Range("A2") = Left(sht3.Range("A2"), 5) & p '从旧表复制过来，表头等有数据，按要求保留、删除
    sht3.Range("H2") = Left(sht3.Range("H2"), 5) & q
    sht3.Range("K2") = Left(sht3.Range("K2"), 6) & r
    sht3.Range("P2") = Left(sht3.Range("P2"), 5)
    sht3.Range("A5:P5000").ClearContents  ' 表中数据全删
    
    On Error GoTo x:
    ' 当输入的工作表名称重复，或者不输名称关掉对话框都会跳错误，
    ' 这时就转到退出sub,新建的工作表仍有效，表名称就是系统自动生成的表名
    
    
    sht3.Name = o
    
x:    Exit Sub
    
End Sub

Private Sub CommandButton1_Click() '输入窗体，按”确定“后取输入的值
o = TextBox1.Value
p = TextBox2.Value
q = TextBox3.Value
r = TextBox4.Value

Unload Me
End Sub
