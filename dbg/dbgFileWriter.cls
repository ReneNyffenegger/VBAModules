'
'  vim: ft=vb
'
'  History
'    0.02   2020-07-30   In dbgWriter_out(): call flush() if flushImmediately_ is true
'
option explicit

implements dbgWriter

private debugFile         as integer
private fileName_         as string
private flushImmediately_ as boolean
'
public sub class_terminate() ' {
' 2019-12-03 Why did I have this `if not` condition?
' 2019-12-03       if not flushImmediately_ then closeFile
    closeFile
end sub ' }

public sub init(fileName as string, optional flushImmediately as boolean = false)
  '
  ' When using dbgFileWriter, it's probably advisable to explicitly
  ' use the VBA `end` statement when done.
  '

    fileName_         = fileName
    flushImmediately_ = flushImmediately

    if debugFile = 0 then
       openFile
    else
       err.raise 999, "dbgFileWriter", "debugFile = " & debugFile
    end if

end sub

public sub dbgWriter_out(txt as string) ' {

    print #debugFile, txt

    if flushImmediately_ then flush

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
