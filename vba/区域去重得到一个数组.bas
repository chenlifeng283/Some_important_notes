Public Function merger_arr(rng As Range)
    '数组去重挺烦的，python只要两三行就解决了
    'targetList = list(set(targetList))
    '可惜离不开VBA，有很多特定操作，VBA的方便性是不能代替的。
    '如录制宏就降低了很多学习记忆成本。
    Dim n As Integer
    Dim mydic As Object
    Dim arr2() As Variant

    '把参数区域数据放入数组
    n = rng.Count
    Dim arr() As Variant
    ReDim arr(1 To n)
    i = 1
    For Each everyRng In rng
        arr(i) = everyRng.Value
        i = i + 1
    Next

    '以arr为字典的键，生成字典，arr中空值不生成
    Set mydic = CreateObject("scripting.dictionary")
    For j = 1 To n
        If arr(j) <> "" Then mydic(arr(j)) = ""
    Next j

    '读取字典key,生成去重后的arr2
    m = mydic.Count
    ReDim arr2(1 To m)
    k = 1
    For Each mydicKey In mydic.keys
        arr2(k) = mydicKey
        k = k + 1
    Next

    '输出
    merger_arr = arr2
    
End Function
