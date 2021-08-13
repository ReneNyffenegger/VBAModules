option explicit

sub test_excelHelpers() ' {

    test_findWorksheet
    test_deleteRange
    test_hyperlink
    test_num_letter

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

private sub test_hyperlink() ' {

    dim ws as worksheet
    set ws = findWorksheet("hyperlinks")

    insertHyperlinkToVBAMacro ws.cells(2,2), "Hyperlink one"                  , "hyperlinked_1"
    insertHyperlinkToVBAMacro ws.cells(3,2), "Hyperlink two (hello world, 42)", "hyperlinked_2", "Hello world", 42
    insertHyperlinkToVBAMacro ws.cells(4,2), "Hyperlink two (foo bar baz, 99)", "hyperlinked_2", "foo bar baz", 99

end sub ' }

private sub test_num_letter_() ' {
  if colLetterToNum("Z" ) <>  26  then msgBox "Expected 26"
  if colLetterToNum("AA") <>  27  then msgBox "Expected 27"
  if colNumToLetter( 26 ) <> "Z"  then msgBox "Expected 'Z'"
  if colNumToLetter( 27 ) <> "AA" then msgBox "Expected 'AA'"
end sub ' }

private sub test_num_letter() ' {

   dim curRefStyle as long
   curRefStyle = application.referenceStyle

   application.referenceStyle = xlA1   : test_num_letter_
   application.referenceStyle = xlR1C1 : test_num_letter_

   application.referenceStyle = curRefStyle

end sub ' }

public function hyperlinked_1() as range ' {
    msgBox "hyperlink one"

  '
  ' The function must return the range that the
  ' hyperlink jumps to
  '
  ' set hyperlinked_1 = sheets("hyperlinks").cells(2,2)
    set hyperlinked_1 = selection
end function ' }

public function hyperlinked_2(txt as string, num as long) as range ' {
    msgBox "hyperlink two, txt = " & txt & ", num = " & num
'   msgBox "hyperlink two, txt = " & txt
    set hyperlinked_2 = selection
end function ' }
