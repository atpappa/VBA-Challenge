sub stock():
    range("I1") = "Ticker"
    range("J1") = "Yearly Change"
    range("K1") = "Percent Change"
    range("L1") = "Total Stock Volume"
    dim yearlyChange as long
    dim loc as long
    loc = 1

    for i = 2 to cells(rows.count,1).end(xlUp).row
    yearlyChange = yearlyChange + (cells(i,6).value - cells(i,3).value)
        if cells(i,1).value <> cells(i+1,1).value then
            loc = loc + 1
            cells(loc,9) = cells(i,1)
        end if
    next i

end sub