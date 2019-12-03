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
' 2019-12-03 Why did I have this `if not` condition?
' 2019-12-03       if not flushImmediately_ then closeFile
    closeFile
end sub ' }

public sub init(fileName as string, optional flushImmediately as boolean = false)
  '
  ' When using dbgFileWriter, it's probably advisable to explicitely
  ' use the VBA `end` statement when done.
  '

    fileName_         = fileName
    flushImmediately_ = flushImmediately

' 2019-12-03    Is it correct to only open the file if `not flushImmediately_`
' 2019-12-03         if not flushImmediately_ then openFile
    if debugFile = 0 then
       openFile
    else
       err.raise 999, "dbgFileWriter", "debugFile = " & debugFile
    end if

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
