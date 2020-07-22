option explicit
'
'   This module is loaded by the moduleLoader.wsf (moduleLoader.bas) test.
'  
'   The following num will be replaced
'   in the test case.
'
const num = 99

sub subInLoadedModule() ' {


    if num <> 42 then
       msgBox "Expected num to be 42!"
    else
       debug.print "num is 42"
    end if

end sub ' }
