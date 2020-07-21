option explicit

sub test_excelHelpers() ' {

    test_findWorksheet
    test_deleteRange

end sub ' }

private sub test_findWorksheet() ' {

    dim ws_foo, ws_bar, ws_baz as worksheet

    set ws_foo = findWorksheet("foo")
    set ws_bar = findWorksheet("bar")
    set ws_baz = findWorksheet("baz")

    ws_foo.cells(1,1) = "foo"
    ws_bar.cells(1,1) = "XXXXX"
    ws_baz.cells(1,1) = "baz"

    set ws_bar = findWorksheet("bar", deleteIfExists := true)
    if ws_bar.cells(1,1) <> "" then ' {
       msgBox "bar: nok"
    end if ' }

    set ws_baz = findWorksheet("baz", deleteIfExists := false)
    if ws_baz.cells(1,1) <> "baz" then ' {
       msgBox "baz: nok"
    end if ' }

    dim sh_nothing as worksheet
    set sh_nothing = collObjectOrNothing(activeWorkbook.sheets, "does not exist")
    if not sh_nothing is nothing then
       msgBox "expected sh_nothing to be nothing"
    end if

end sub ' }

private sub test_deleteRange() ' {

    dim ws_rng as worksheet
    set ws_rng = findWorksheet("rng")

    with ws_rng
     
         with .range( .cells(3,2), .cells(6, 4) ) ' {
             .value = "A"
             .interior.color = rgb(255, 135, 40)

             .name = "A"
         end with ' }

         with .range( .cells(6,3), .cells(7, 5) ) ' {
             .value = "B"
             .interior.color = rgb( 40, 180,220)

             .name = "B"
         end with ' }

        deleteRange "A"
        deleteRange "thisRangeDoesNotExist"

        with .cells(6, 4)

            if .text <> "" then
               msgBox "test_deleteRange failed (1)"
            end if

            if .interior.Color <> rgb(255, 255, 255) then
               msgBox "test_deleteRange failed (2)"
            end if

        end with

 
    end with

end sub ' }
