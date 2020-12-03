Sub 多表多文件合并为多表一文件()

Dim FileArray

Dim X As Integer
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Application.ScreenUpdating = False

On Error GoTo Y

FileArray = Application.GetOpenFilename(FileFilter:="Microsoft Excel文件(*.xls),*.xls", MultiSelect:=True, Title:="合并工作薄")

X = 1

While X <= UBound(FileArray)
    startRow = ThisWorkbook.Worksheets("发票信息").Range("c50000").End(xlUp).Row + 1
    startRow2 = ThisWorkbook.Worksheets("货物信息").Range("c50000").End(xlUp).Row + 1
    
    
    Workbooks.Open Filename:=FileArray(X)
    Set wb = ActiveWorkbook
    Set ws1 = ActiveWorkbook.Worksheets("发票信息")
    Set ws2 = ActiveWorkbook.Worksheets("货物信息")
    ws1.Select
    endrow = Range("c50000").End(xlUp).Row
    If VBA.InStr(1, Range("a1").Value, "普通") Then
        Range("a3:z" & endrow).Copy ThisWorkbook.Worksheets("发票信息").Range("A" & startRow)
        ThisWorkbook.Worksheets("发票信息").Activate
        ThisWorkbook.Worksheets("发票信息").Range(Cells(startRow, 5), Cells(startRow + endrow - 3, 5)).Value = "普票"
    Else:
        Range("a3:z" & endrow).Copy ThisWorkbook.Worksheets("发票信息").Range("A" & startRow)
    End If
    
    wb.Activate
    ws2.Select
    endrow2 = Range("c50000").End(xlUp).Row
    If VBA.InStr(1, Range("a1").Value, "普通") Then
        Range("a3:z" & endrow2).Copy ThisWorkbook.Worksheets("货物信息").Range("A" & startRow2)
        ThisWorkbook.Worksheets("货物信息").Activate
        ThisWorkbook.Worksheets("货物信息").Range(Cells(startRow2, 4), Cells(startRow2 + endrow2 - 3, 4)).Value = "普票"
    Else
        Range("a3:z" & endrow2).Copy ThisWorkbook.Worksheets("货物信息").Range("A" & startRow2)
    End If
    wb.Close


'Sheets().Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

X = X + 1

Wend
MsgBox "导入成功"

ExitHandler:

Application.ScreenUpdating = True

Exit Sub

errhadler:

   MsgBox Err.Description

Y:
End Sub
