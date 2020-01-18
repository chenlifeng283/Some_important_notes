'函数要写在“模块”下
'判断某区域（arr)里与某一单元格（single_arr)底色相同的单元格个数
Function color_count(arr as range,single_arr as range)
  Application.volatile True '把函数设为易失性函数，所引用的单元格内容发生变化就会自动重算，调试效果好像有问题。
  ' 条件是单元格重算，底色改变不算是单元格重算，因此本函数调试无效。
  dim rng as range
  
  For each rng in arr
  
    if rng.interior.color = single_arr.interior.color then
    color_count = color_count + 1 '自定义函数返回的结果，要保存到过程名称（color_count)中。
    end if
    
  next rng
End Function
