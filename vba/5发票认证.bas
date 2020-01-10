# 涉及到的知识点：1.怎么通过对话框选择文件，获取文件路径，可以引申为选择文件夹
                2.跨工作薄粘贴要注意“随时特用的工作表”，以及Public变量申明
                3.选择性粘贴为数值
                4.控制不准改工作表
                5.有公式的单元格无法触发worksheet_change事件，需要通过worksheet_calculate事件
        
Option Explicit

Public FileDialogObject As Object
Public paths As Object
Public p
Public wb As Workbook
Public j

Sub FilePicker()
    ' 选择文件对话框
    Set FileDialogObject = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialogObject
        .Title = "选择从‘发票认证系统’中导出的发票清单文件"
        .InitialFileName = "C:\Users\admin\Desktop"

    End With
    FileDialogObject.Show
    Set paths = FileDialogObject.SelectedItems

    p = paths.Item(1) '取选择的文件的路径字符串
    Workbooks.Open Filename:=p '按路径打开文件
    Set wb = ActiveWorkbook  ' wb表示打开的文件
    j = Application.CountA(wb.Worksheets(Sheet1).Range("A:A")) '清单你有几行
    wb.Worksheets(Sheet1).Range(Cells(2, 1), Cells(j, 10)).Select '选择
    Selection.Copy ThisWorkbook.Worksheets("网上下载清单").Range("A2") '复制到模板表

    ThisWorkbook.Worksheets("手工输入发票清单").Activate

End Sub


Sub closeFinal()

    ThisWorkbook.Worksheets("网上下载清单").Activate '在用表前要先激活，不然点到其他表再运行会跳错误
    Range("A1").Select

    'ThisWorkbook.Worksheets("网上下载清单").
    Range(Cells(2, 13), Cells(j, 13)).Select  '
    Selection.Copy
    wb.Worksheets(Sheet1).Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False '选择性粘贴

    wb.Close savechanges:=True

End Sub

' 下面是通过一些“事件”来优化程序

 ' 防止改工作表名称                
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Me.Name <> "手工输入发票清单" Then Me.Name = "手工输入发票清单"
End Sub

                    
' 根据H5的计算结果来判断有没有输入错误                    
Private Sub Worksheet_Calculate()
Dim i As Integer
    Range("C2:C1000").Interior.ColorIndex = xlNone
    '每次输入都会触发这个程序,初始化C列的颜色,因为后面代码有错误标红后,
    ' 想改正后去掉标红,发现找不到合适的位置写这条代码,索性每次回复一遍无底色,
    ' 再通过下面的代码"有错标红",肯定有更好的办法
    
     If Range("H5") <> 0 Then  '每次H5重算后，就判断
        For i = 2 To 1000
            If Cells(i, 3) <> "" And Cells(i, 5) <> 0 Then
            '表格特点：C列为空时，是不会有错的，如果C列填入数据，E列还不是零，那肯定是输错了。
                Cells(i, 3).Interior.Color = RGB(255, 0, 0)
                ' Else:  Cells(i, 3).Interior.ColorIndex = xlNone
            End If
        Next
     Else:
         Exit Sub ' H5为零时就退出程序
     End If    
End Sub





