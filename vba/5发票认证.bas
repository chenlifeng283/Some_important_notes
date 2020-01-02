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
