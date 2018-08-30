'
'  vim: ft=vb
'
option explicit

implements dbgWriter

private debugFile         as integer
private fileName_         as string
private flushImmediately_ as boolean
'
public sub class_terminate() ' {
    if not flushImmediatly_ then closeFile
end sub ' }

public sub init(fileName as string, optional flushImmediately as boolean = false)
    fileName_         = fileName
    flushImmediately_ = flushImmediately

    if not flushImmediately_ then openFile
end sub

public sub dbgWriter_out(txt as string) ' {

    if flushImmediately_ then openFile

    print #debugFile, txt

    if flushImmediately_ then   closeFile

end sub ' }

public sub flush() ' {
    closeFile
    openFile
end sub ' }

private sub openFile() ' {

    debugFile = freeFile()
    open fileName_ for append as #debugFile

end sub ' }
'
private sub closeFile() ' {
    if debugFile <> 0 then
       close #debugFile
    end if
    debugFile = 0
end sub ' }
