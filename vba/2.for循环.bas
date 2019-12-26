sub toUSD()
' 把第六列的人民币价转换为USD价
    dim rate
    rate = cells(8,6)
    
    for i =11 to 20 step 1
        cells(i,6) = cells(i,6)/rate
    next i ' i 可以省略，嵌套多了，不写分不清，习惯。

    cells(7,6) = "USD"
end sub
