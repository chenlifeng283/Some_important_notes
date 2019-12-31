' 1.定义变量，并存储数据，将数据写入到指定单元格
Sub 数据变量()
  Dim IntCount
  IntCount = 3000
  Range("A1").Value = IntCount
End Sub

'2.定义对象变量
Sub 对象变量()
    Dim sht As Worksheet
    Set sht = ActiveSheet
    Sht.range("A1").value = " I'm learning..."
 End sub
 
 '3.定义模块级变量
 Dim a As String
 Private b As String
 
 Sub 合并文本()
    a = "learing1..."
    b = :learning2..."
    MsgBox a & b
  End sub
  
'4.声明动态数组
  Sub Test()
      Dim a As Interger
      a = Application.WorkshhetFunction.CountA(Range("A:A"))
      Dim arr() As String '定义一个Sting类型的动态数组
      Redim arr(1 to a) '重新定义数组arr的大小
   End Sub
   
'5.通过单元格数据直接创建数组
   Sub RngArr_1()
      Dim arr As Variant
      arr= Range("A1:c3").Value
      Range("E1:G3").Value = arr
    End Sub
    
    Sub RngArr_2()
        Dim arr as Variant
        arr = Range("A1:C3").Value
        MsgBox arr(2,3)  '显示arr中第2行的第3个数据
     End Sub
     
'6.用Split函数创建数组
Sub SplitTest()
    Dim a As Variant
    arr = Split("one,two,three,four",",")
    Msgbox "arr数组中的第2个元素是：" & arr(1)
End sub

'7.用Array函数创建数组
Sub ArrayTest()
    Dim a As variant
    arr = Array("one","two","three")
    Msgbox "arr数组中的第2个元素是：" & arr(1)
    MsgBox "数组的最大索引号是：" & UBound(arr)
    MsgBOx "数组的最小索引号是：" & LBound(arr)
End Sub

'8.用Join函数将一维数组合并
Sub JoinTest()
    Dim arr As Variant,txt As String
    arr = Array(0,1,2,3,4,5)
    txt = Join(arr,"@")  '与split功能相反
    Msgbox txt
End Sub     

'9.求数组包含个数
Sub ArrayTest()
    Dim arr as Variant
    arr = Array(1,2,3,4,5)
    Dim a As Integer,b As Integer
    a = UBound(arr)
    b = LBound(arr)
    MsgBox "数组包含的元素个数是：" & a-b+1
 End Sub
 
 Sub RngArr()
    Dim arr As Variant
    arr = Range("A1:C3").Value
    Dim a As Integer,b As Integer
    a = UBound(arr,1) '求数组第一维的最大索引号
    b = Lbound(arr,1)
    Dim c As Interger, d As Integer
    c = Ubound(arr,2)
    d = LBound(arr,2)
    MsgBox "数组包含的元素个数是：" & (a-b+1)*(a-b+1)
 End Sub
 
 '10.将数组中保存的数据写入单元格
 Sub ArrToRng1()
    Dim arr as Variant
    arr = Array(1,2,3,4,5)
    Range("A1:A9").Value = Application.WorksheetFunction.Transpose(arr)
 End Sub
 
 Sub ArrToRng2（）
    Dim arr(1 to 2,1 to 3) As String
    arr(1,1) = 1
    arr(1,2) = "clf"
    arr(1,3) = "male"
    arr(2,1) = 2
    arr(2,2) = "blf"
    arr(2,3) = "female"
    Range("A1:C2").Value = arr
 End Sub
 
 11.声明常量
 Sub ArrayTest()
    Const P As Single = 3.14
 End Sub
 
 12.用IF语句判断成绩等级
 Sub Test_1()
    If Range("B2").value >= 90 Then
        Range("C2").Value = "优秀"
    Else
        If Range("B2").Value >= 80 Then
            Range("C2").Value = "良好"
        Else
            If Range("B2").Value >= 60 Then
                Range("C2").Value = "及格"
            Else
                Range("C2").Value = "不及格"
            End if
         End if
      End if 
   End Sub
   
'13. 用Exit for 终止循环
Sub ShtAdd()
    Dim i As Byte
    for i = 1 To 5 Step 1
        Worksheets.Add
        exit for
     Next i
 End Sub
 
    
 
 
 

