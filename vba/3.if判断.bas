sub evaluae()
    dim score

    score = (cells(4,6))+cells(5,6))/3
    cells(7,6) = score
    if score >= 60 then
        cells(8,6)="及格"
    else
        cells(8,6)="不及格"
    end if

end sub

'''
if *** then
    ***
elseif *** then
    ***
else
    ***
end if
'''