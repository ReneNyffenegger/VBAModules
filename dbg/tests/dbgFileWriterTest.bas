option explicit

global dbg_ as new dbg

sub main() ' {

    dim fileWriter as new dbgFileWriter
    fileWriter.init(environ("TEMP") & "\dbgFileWriterTest.txt")

    call dbg_.init(fileWriter)

    dbg_.text("dbg_ is initialized")

    dbg_.text("func_one returned " & func_one(42))


end sub ' }

function func_one(num as long) as long ' {
    dbg_.indent("func_one, num = " & num)

    if num > 40 then
       dbg_.text("num > 40, call func_two")
       func_one = func_two(num-6)
    elseif num > 20 then
       dbg_.text("num > 20, call func_one")
       func_one = func_three(num-12)
    else
       dbg_.text("add 5")
       func_one = num + 5
    end if

    dbg_.text("returning " & func_one)

    dbg_.dedent

end function ' }

function func_two(num as long) as long ' {
    dbg_.indent("func_two, num = " & num)

    if num > 20 then
       dbg_.text("num > 20, call func_three")
       func_two = func_three(num-14)
    elseif num > 10 then
       dbg_.text("num > 10, call func_one")
       func_two = func_one(num+7)
    else
       dbg_.text("adding 6")
       func_two = num + 6
    end if

    dbg_.text("returning " & func_two)

    dbg_.dedent

end function ' }

function func_three(num as long) as long ' {
    dbg_.indent("func_three, num = " & num)

    if num > 40 then
       dbg_.text("num > 40, call func_two")
       func_three = func_two(num-8)
    elseif num > 20 then
       dbg_.text("num > 20, call func_three")
       func_three = func_one(num+3)
    else
       dbg_.text("add 4")
       func_three= num + 4
    end if

    dbg_.text("returning " & func_three)

    dbg_.dedent

end function ' }
