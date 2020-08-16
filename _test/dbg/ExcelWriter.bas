option explicit

dim dbg_ as dbg

sub main() ' {

    set dbg_ = new dbg

    dim excelWriter as new dbgExcelWriter
    excelWriter.init activeWorkbook, "dbg"
    dbg_.init excelWriter

    dbg_.text "started"

    F1
    F2

end sub ' }

sub F1() ' {
    dbg_.indent  "F1"
    dbg_.text "in F1"
    F2
    dbg_.dedent
end sub ' }

sub F2() ' {
    dbg_.indent  "F2"
    dbg_.text "in F2"

    dbg_.dedent
end sub ' }
