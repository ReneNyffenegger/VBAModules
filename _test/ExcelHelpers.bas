option explicit

sub test_excelHelpers() ' {

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
