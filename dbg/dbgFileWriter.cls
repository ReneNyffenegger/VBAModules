'
'  vim: ft=vb
'
'  History
'    0.02   2020-07-30   In dbgWriter_out(): call flush() if flushImmediately_ is true
'    0.03   2021-07-20   Fix mess with flushImmediately_
'
option explicit

implements dbgWriter

private debugFile         as integer
private fileName_         as string
private flushImmediately_ as boolean
'
public sub class_terminate() ' {
   if not flushImmediately_ then
      closeFile
   end if
end sub ' }

public sub init(fileName as string, optional flushImmediately as boolean = false)
  '
  ' When using dbgFileWriter, it's probably advisable to explicitely
  ' use the VBA `end` statement when done.
  '

    fileName_         = fileName
    flushImmediately_ = flushImmediately

    if not flushImmediately_ then
       openFile
    end if

end sub

public sub dbgWriter_out(txt as string) ' {

    if flushImmediately_ then
       openFile
    end if

    print #debugFile, txt

    if flushImmediately_ then
       closeFile
    end if

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
