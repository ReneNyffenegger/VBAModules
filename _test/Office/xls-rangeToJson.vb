option explicit

sub main() ' {

    cells(2,2) = "num" : cells(2, 3) = "txt": cells(2, 4) = "dat"                 : cells(2,5) = "msc"
    cells(3,2) =    1  : cells(3, 3) = "two": cells(3, 4) = #2020-08-28#          :
    cells(4,2) =  2.2  : cells(4, 3) =    2 : cells(4, 4) = #2021-08-28 22:23:24# : cells(4,5) = ""

    dim jsn as string
    jsn = excelRangeToJson(range(cells(2,2), cells(4,5)))

    if jsn <> "[[""num"",""txt"",""dat"",""msc""],[1,""two"",""2020-08-28T00:00:00"",null],[2.2,2,""2021-08-28T22:23:24"",null]]" then
       msgBox "jsn = " & jsn
    end if

end sub ' }
