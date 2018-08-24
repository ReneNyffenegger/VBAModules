'
'  vim: ft=vb
'
option explicit

implements dbgWriter

private debugFile   as integer
private fileName_   as string
'
public sub class_terminate() ' {
    closeFile
end sub ' }

public sub init(fileName as string)
    fileName_ = fileName
    openFile
end sub

public sub dbgWriter_out(txt as string) ' {
  ' openFile
    print #debugFile, txt
   'closeFile
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
