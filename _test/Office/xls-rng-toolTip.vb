option explicit

sub main() ' {

   cells(2,2) = "Foo" : excelRangeToolTip cells(2,3), "Lorem ipsum"           , "Lorem ipsum dolor sit amet,"
   cells(3,2) = "Bar" : excelRangeToolTip cells(3,3), "consectetur adipiscing", "elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua"
   cells(4,2) = "Baz" : excelRangeToolTip cells(4,3), "Ut enim"               , "ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat"

end sub ' }
