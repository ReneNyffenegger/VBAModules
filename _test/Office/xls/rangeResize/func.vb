option explicit

sub test_excelRangeResize() ' {

    dim rng as range
    set rng = range(cells(3,4), cells(4,6))

    rng.interior.color = rgb(180, 210, 255)

    dim rngResized as range
    set rngResized = excelRangeResize(rng, leftRel := 2, rightRel := 3, topRel := -1)
    rngResized.borderAround xlDash, xlMedium, color := rgb(290, 100, 255)

    range(columns(1), columns(10)).columnWidth = 2

    cells(10,1).select

end sub ' }
