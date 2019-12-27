'判断某区域（arr)里与某一单元格（single_arr)底色相同的单元格个数
Function color_count(arr as range,single_arr as range)
  dim rng as range
  
  For each rng in arr
  
    if rng.interior.color = single_arr.interior.color then
    color_count = color_count + 1 '自定义函数返回的结果，要保存到过程名称（color_count)中。
    end if
    
  next rng
End Function
