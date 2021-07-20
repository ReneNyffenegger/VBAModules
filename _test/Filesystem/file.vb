option explicit

sub main() ' {

    dim file_foo as file
    dim file_bar as file
    dim file_baz as file

    set file_foo = new file
    set file_bar = new file
    set file_baz = new file

    file_foo.open_ environ("temp") & "\filetest.foo.txt"
    file_bar.open_ environ("temp") & "\filetest.bar.txt"
    file_baz.open_ environ("temp") & "\filetest.baz.txt"

    file_foo.print_      "This "
    file_foo.print_      "is "
    file_foo.print_      "the "
    file_foo.print_      "first "
    file_foo.printWithNL "line. "

    file_bar.printWithNL "This is the bar file."
    file_baz.printWithNL "This is the baz file."

    file_foo.printWithNL "Second line"
    file_bar.printWithNL "Second line"
    file_baz.printWithNL "Second line"

    file_foo.close_
    file_bar.close_
    file_baz.close_

end sub ' }
