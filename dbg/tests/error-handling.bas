option explicit

global dbg_ as new dbg

sub main() ' {
  on error goto err_

    dim fileWriter as new dbgFileWriter
    fileWriter.init(environ("TEMP") & "\error-handling.txt")
    call dbg_.init(fileWriter)
    dbg_.indent("main")

    sub_one
    sub_two
    sub_three

    dbg_.dedent
    exit sub

  err_:
    dbg_.dedent

    msgBox("Error: "& err.description)

end sub ' }

sub sub_one() ' {
  on error goto err_
    dbg_.indent("sub_one")

    dbg_.dedent
    exit sub

  err_:
    dbg_.unhandledError
end sub ' }

sub sub_two() ' {
  on error goto err_
    dbg_.indent("sub_two")

    sub_sub

    dbg_.dedent
    exit sub

  err_:
    dbg_.unhandledError
end sub ' }

sub sub_three() ' {
  on error goto err_
    dbg_.indent("sub_three")

    dbg_.dedent
    exit sub

  err_:
    dbg_.unhandledError
end sub ' }

sub sub_sub() ' {
  on error goto err_
    dbg_.indent("sub_sub")

    sub_sub_sub ( 6)
    sub_sub_sub ( 0) ' boom
    sub_sub_sub (15)

    dbg_.dedent
    exit sub

  err_:
    dbg_.unhandledError
end sub ' }

sub sub_sub_sub(p as long) ' {
  on error goto err_
    dbg_.indent("sub_sub_sub, p = " & p)

    dbg_.text "30 / " & p & " = " & (5/p)

    dbg_.dedent
    exit sub

  err_:
    dbg_.unhandledError
end sub ' }
