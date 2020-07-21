option explicit

sub test_Text() ' {

    test_rpad_lpad 

end sub ' }

sub test_rpad_lpad() ' {

    if rpad("abc", 10, ".") <> "abc......." then ' {
       msgBox "expected: abc......."
    else
       debug.print("rpad: ok")
    end if ' }

    if lpad("xyz", 10, ".") <> ".......xyz" then ' {
       msgBox "expected: .......xyz"
    else
       debug.print("lpad: ok")
    end if ' }

end sub
