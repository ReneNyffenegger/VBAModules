'
' vi: ft=basic
'

option explicit

private fileNumber as integer

public sub open_(name_ as string) ' {

    if fileNumber <> 0 then ' {
       msgBox("fileNumber already set, name_ = " & name_)
       exit sub
    end if ' }

    fileNumber = freeFile()
    open name_ for output as #fileNumber

end sub ' }

public sub print_(txt as string) ' {
    print# fileNumber, txt; ' Semicolon to prevent printing of new line
end sub ' }

public sub printWithNL(txt as string) ' {
    print# fileNumber, txt  ' No semicolon to prevent printing of new line
end sub ' }

public sub close_ ' {

    close #fileNumber
    fileNumber = 0

end sub ' }

private sub class_initialize()
end sub

private sub class_terminate()
end sub
