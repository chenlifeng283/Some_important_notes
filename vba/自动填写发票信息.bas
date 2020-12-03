'修正了品名栏不能为空的要求
'提了3倍速度
Public start_row As Integer
Public TS2 As Integer
Public notFound As Integer

Sub one_step()
t1 = Timer
Call 提示重复输入2
If TS2 > 1 Then
    MsgBox "T列有重复或错误的发票号码，已标记颜色"""
    Exit Sub
End If
Call 填表2

Call 清理代开企业数据2
If notFound > 1 Then
    MsgBox "T列有找不到发票信息的，已标颜色"
End If
Call 发票号码提示重复输入2

t2 = Timer
'MsgBox t2 - t1

End Sub

Sub 填表2()
Dim sht1 As Worksheet
Dim sht2 As Worksheet
Dim sht3 As Worksheet
Set sht1 = ActiveSheet
Set sht2 = Worksheets("发票信息")
Set sht3 = Worksheets("货物信息")
Dim arr() As Variant '货物信息表中10列数据
Dim arr3() As Variant '存入数据，最后一次性写到明细表里
Dim str1 As String

n = sht3.Range("C50000").End(xlUp).Row
ReDim arr(3 To n, 1 To 10)
arr = sht3.Range("C3:L" & n)

'很难确定arr3的准确个数，给最大值，保证数据紧临写入，后面全是空值就行了，
'实际也可写个循环去掉空值，但这里不影响数据，没必要
ReDim arr3(1 To n, 1 To 8)

Application.ScreenUpdating = False
'把辅助列格式初始化
ActiveSheet.Range("t5:T500").Interior.ColorIndex = xlNone
ActiveSheet.Range("c5:c500").Interior.ColorIndex = xlNone

'判断T列输入的发票号是否在arr里，用来判断输入的发票有没有找到
m = ActiveSheet.Range("t5000").End(xlUp).Row
X = 1
notFound = 1
For i = 5 To m
    
    If (sht3.Range("C3:C" & n).Find(ActiveSheet.Cells(i, 20)) Is Nothing Or Len(ActiveSheet.Cells(i, 20)) <> 8) And ActiveSheet.Cells(i, 20) <> "" Then
        ActiveSheet.Cells(i, 20).Interior.ColorIndex = 35
        notFound = notFound + 1
    End If
    
    For a = LBound(arr) To UBound(arr)
        If arr(a, 1) = ActiveSheet.Cells(i, 20) Then
            str1 = arr(a, 3)
            
            pm = split_str(str1) '调用自定义函数，得到符合要求的品名
            
            If pm <> "" Then '品名为“详见销货清单”的，返回值为空，这一行数据是不要的，不用存入数组
                arr3(X, 1) = arr(a, 1)
                arr3(X, 2) = ""
                arr3(X, 3) = Trim(pm)
                arr3(X, 4) = arr(a, 5)
                arr3(X, 5) = arr(a, 6)
                If arr(a, 2) = "普票" Then
                    arr3(X, 6) = arr(a, 8) * 1 + arr(a, 10) * 1 '不含税金额
                    arr3(X, 8) = ""
                Else:
                    arr3(X, 6) = arr(a, 8) * 1
                    arr3(X, 8) = arr(a, 10) * 1
                End If
                arr3(X, 7) = arr(a, 9) * 1
                
                X = X + 1  'X放哪影响数据顺序，调试得结果
            End If
        End If
    Next
Next
ActiveSheet.Range("C:C").NumberFormatLocal = "@" '定义成文本格式，不定义要出错，发票号码很容易出现格式问题

'定位没有数据的一行，开始连续写入数据
'不能用union连接不连续区域然后用find查找，
'以下两个区域分别定位，取较大数就可以了

Dim r1 As Range
Dim r2 As Range
Set r1 = Range("A4:J5000")
Set r2 = Range("L4:r5000")
mm = r1.Find("*", searchdirection:=xlPrevious, searchorder:=xlRows).Row + 1
nn = r2.Find("*", searchdirection:=xlPrevious, searchorder:=xlRows).Row + 1
If mm > nn Then
    start_row = mm
Else
    start_row = nn
End If

Cells(start_row, 3).Select
Selection.Resize(n, 8) = arr3

Application.ScreenUpdating = True

End Sub
Sub 清理代开企业数据2()

Dim sht As Worksheet
Dim sht1 As Worksheet
Dim mydic As Object
Set mydic = CreateObject("scripting.dictionary")
Dim s_col_value As String


Set sht = ActiveSheet
Set sht1 = Worksheets("发票信息")
n = Range("C50000").End(xlUp).Row
'Debug.Print n
'形成一个字典，以c列发票号码为key,以发票信息表里的备注栏列，经正则表达示匹配后的企业名为item

For i = start_row To n
    If sht.Cells(i, 3) <> "" Then
         sht.Cells(i, 1) = Application.WorksheetFunction.VLookup(sht.Cells(i, 3), sht1.Range("C:s"), 2, 0)
        mykey = sht.Cells(i, 3)
        '用vlookup()查到发票号对应的“备注”数据，如果包含“代开企业”说明是代开企业
        '若不包含，正常企业，直接用vlookup查数据就行了
        '为结构明确，写了reg()函数生成所需数据
        
        s_col_value = Application.WorksheetFunction.VLookup(sht.Cells(i, 3), sht1.Range("C:s"), 17, 0)
            If VBA.InStr(1, s_col_value, "代开企业") Then
                myitem = reg2(s_col_value)
            Else
                myitem = Application.WorksheetFunction.VLookup(sht.Cells(i, 3), sht1.Range("C:s"), 5, 0)
            End If
        '以发票号为key,单位名称为item,往mydic写入数据
        '实际循环下来，同一组数据实际写入多次，但因为字典key唯一性的特点，最终会形成一个号码对应一个企业的效果
        mydic(mykey) = myitem
        '下面这句，可以等字典形成后，再写循环写入，更好理解，但这里效果合并了。
        sht.Cells(i, 4).Value = mydic(sht.Cells(i, 3).Value)
    End If
Next

'For Each j In mydic.items
'    Debug.Print j
'Next
'

'同一发票，同一企业名称，只显示一行
'因为按程序特点，所有重复的数据是顺序排在一起的
'只要计算出个数，扩大range范围，直接合并就行了，再全部取消合并

'如果第二批发票输入有这类发票的重复，那这个“个数”就会出错，数据会有错行
Application.DisplayAlerts = False
For i = start_row To n
    If sht.Cells(i, 3) <> "" Then
        m = Application.WorksheetFunction.CountIf(sht.Range(Cells(start_row, 3), Cells(n, 3)), Cells(i, 3))
        sht.Range(Cells(i, 3), Cells(i + m - 1, 3)).Merge
        sht.Range(Cells(i, 4), Cells(i + m - 1, 4)).Merge
        sht.Range(Cells(i, 1), Cells(i + m - 1, 1)).Merge
            
    End If
Next

sht.Range("c5:c" & n).UnMerge
sht.Range("d5:d" & n).UnMerge
sht.Range("a5:a" & n).UnMerge
sht.Range("a5:d" & n).Borders.LineStyle = xlContinuous
sht.Range("a5:d" & n).WrapText = False
    
Application.DisplayAlerts = True


End Sub
Sub 提示重复输入2()
Dim rng As Range
Dim arr() As Variant

n = ActiveSheet.Range("t5000").End(xlUp).Row
If n = 4 Then Exit Sub
ReDim arr(5 To n)
'arr = Range(Cells(5, 20), Cells(n, 200))

'把T列发票号都放入数组，然后每个数据依次与数组去筛选，Uboound(新数组）>0的有两个以上
Q = 5
For Each rng In Range("t5:t" & n)
    arr(Q) = rng
Q = Q + 1
Next

'全局变量，进入有重复的判断后，值会大于1,用来在总过程中判断是否要跳出msgbox
TS2 = 1
For i = 5 To n
    If UBound(Filter(arr, Cells(i, 20).Value)) > 0 And Cells(i, 20) <> "" Then
        Cells(i, 20).Interior.ColorIndex = 45
        TS2 = TS2 + 1
    End If
Next


End Sub
Sub 链接2()
Dim sht As Worksheet
Dim arr() As Variant
Dim arr2(1 To 7) As Variant

Application.ScreenUpdating = False
Range("B2:i10000") = ""
n = Worksheets.Count
ReDim arr(1 To n - 3)
i = 1

For Each sht In Worksheets
    If sht.Name <> "目录" And sht.Name <> "发票信息" And sht.Name <> "货物信息" Then
        arr(i) = sht.Name
        i = i + 1
    End If
    
Next

Range("b2:B" & n - 2) = Application.WorksheetFunction.Transpose(arr)

For Each everyRng In Range("B2:b" & n - 2)
    everyRng.Select
    Selection.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & everyRng.Value & "'!A1", ScreenTip:="", TextToDisplay:=everyRng.Value
    Union(Worksheets(everyRng.Value).Range("j4:k4"), Worksheets(everyRng.Value).Range("M4:Q4")).Copy Selection.Offset(0, 1)
Next

ActiveSheet.Range("c2:i5000").Interior.ColorIndex = xlNone
ActiveSheet.Range("c2:i5000").Borders.LineStyle = xlContinuous

Application.ScreenUpdating = True
End Sub
Sub 新建表格2()
Dim sht As Worksheet
ActiveSheet.Copy after:=ActiveSheet
ActiveSheet.Range("t5:T500").Interior.ColorIndex = xlNone
ActiveSheet.Range("a5:j10000").Value = ""
ActiveSheet.Range("l5:t10000").Value = ""
With ActiveSheet
    .Range("A2").Value = "工程名称:"
    .Range("H2").Value = "合同金额："
    .Range("k2").Value = "项目承包人："
    .Range("P2").Value = "合同日期："
End With
    


End Sub

Function reg2(str1 As String)

Dim myreg As Object
Dim mymatches As Object

Set myreg = CreateObject("vbscript.regexp")
'以“代开企业名称”为起点，以“司部”等字结尾，有局限
'如，公司名中间就出现“行所”等字样，就会取不到完整名称
'如，直接是个人姓名，就取不到了，没特点，写不好完全匹配的正则表达式

myreg.Pattern = "代开企业名称[:：](.+[司部厂行店场户站队处所心校\s]*)"
myreg.Global = True
Set mymatches = myreg.Execute(str1)
If mymatches.Count = 0 Then
    reg2 = "代开"
Else
    reg2 = "(代开)" & mymatches(0).submatches(0)
End If

'On Error GoTo X
'reg = mymatches(0).submatches(0)
'Exit Function
'X: reg = "代开"
End Function
Function split_str(whatstr As String)
Dim arr As Variant

    If VBA.InStr(1, whatstr, "*") Then
        arr = VBA.Split(whatstr, "*")
    ElseIf whatstr = "(详见销货清单)" Then
        Exit Function
    Else
        arr = Array(whatstr)
    End If
    
    split_str = arr(UBound(arr))
End Function
Sub merge_same()
Dim arr As Variant
arr = Range("t5:t500")
Set rng = Selection
start_row = rng.Row '起始行号
rng_count = rng.Count '有几行
If rng_count = 1 Then Exit Sub
Cells(start_row, 7) = Application.WorksheetFunction.Sum(Range(Cells(start_row, 7), Cells(start_row + rng_count - 1, 7)))
Cells(start_row, 8) = Application.WorksheetFunction.Sum(Range(Cells(start_row, 8), Cells(start_row + rng_count - 1, 8)))
Cells(start_row, 10) = Application.WorksheetFunction.Sum(Range(Cells(start_row, 10), Cells(start_row + rng_count - 1, 10)))
Range(Cells(start_row + 1, 7), Cells(start_row + rng_count - 1, 7)).EntireRow.Delete
Range("t5:t500") = arr

End Sub
Sub 发票号码提示重复输入2()
Dim rng As Range
Dim arr() As Variant

n = ActiveSheet.Range("c15000").End(xlUp).Row
If n = 3 Then Exit Sub

ReDim arr(5 To n)
'arr = Range(Cells(5, 20), Cells(n, 200))

'把T列发票号都放入数组，然后每个数据依次与数组去筛选，Uboound(新数组）>0的有两个以上
Q = 5
For Each rng In Range("c5:c" & n)
    arr(Q) = rng
Q = Q + 1
Next

'全局变量，进入有重复的判断后，值会大于1,用来在总过程中判断是否要跳出msgbox
TS = 1
For i = 5 To n
    If UBound(Filter(arr, Cells(i, 3).Value)) > 0 And Cells(i, 3) <> "" Then
        Cells(i, 3).Interior.ColorIndex = 45
        TS = TS + 1
    End If
Next

If TS > 1 Then MsgBox "发票号码列有重复输入,请核对"

End Sub
