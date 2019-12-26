option explicit '强制申明

submathtest ()
    dim r1,s,v '申明
    const pi=3.14 '定义常量,不能给pi二次赋值

    r1 = cells(4,3)
    s = pi*r1*r1
    v = 4/3*pi*r1*r1

    cells(4,4) = s
    cells(4,5) = v
end sub
