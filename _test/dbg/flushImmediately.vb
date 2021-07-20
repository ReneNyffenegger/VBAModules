option explicit

dim dbg_ as dbg

sub main(currentDir as variant) ' {

    set dbg_ = new dbg 
     
    dim fileWriter as new dbgFileWriter
    fileWriter.init currentDir & "dbg-out_" & format(now, "yyyy-mm-dd_hhnn") & ".txt", flushImmediately := true
    dbg_.init fileWriter

    dbg_.text "started"

    A

end sub ' }

sub A() ' {
    dbg_.indent "A"
    dbg_.text "in A"
    B
    dbg_.dedent
end sub ' }

sub B() ' {
    dbg_.indent "B"
    dbg_.text "in B"
    dbg_.dedent
end sub ' }
