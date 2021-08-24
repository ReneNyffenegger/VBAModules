option explicit

sub main() ' {

    test_subtraction range(cells( 2,2), cells(7 ,6)), range(cells(3,3), cells(6,5))
    test_subtraction range(cells(12,2), cells(17,6)), range(rows(12), rows(13))
    test_subtraction range(cells(22,2), cells(27,6)), range(cells(23,3), cells(28,7))

    test_subtraction excel.application.union(range( cells(30,2), cells(31,3) ) ,    _
                                             range( cells(33,2), cells(34,3) ) ,    _
                                             range( cells(30,6), cells(31,7) ) ,    _
                                             range( cells(33,6), cells(34,7) ) ) ,  _
                                             range( cells(31,3), cells(33,6))


end sub ' }


sub test_subtraction(rng_1 as range, rng_2 as range) ' {

    rng_1.borderAround xlContinuous, xlMedium, color := rgb(240, 10,  40)
    rng_2.borderAround xlContinuous, xlMedium, color := rgb( 30, 10, 200)

    dim rngRes as range
    set rngRes = excelRangeSubtract(rng_1, rng_2)

    rngRes.interior.color = rgb(255, 240, 190)

    debug.print "rngRes.areas.count = " & rngRes.areas.count

end sub ' }
