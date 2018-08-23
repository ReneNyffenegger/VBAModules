'
'  vim: ft=vb
'
option explicit

private indent_     as integer
private debugFile   as integer
private fileName_   as string

public sub class_initialize() ' {
    debugFile = 0
end sub ' }

public sub class_terminate() ' {
    if debugFile <> 0 then
       closeFile
    end if
end sub ' }

public sub init(fileName as string)
    indent_   = 0
    fileName_ = fileName
end sub

public sub text(txt as string) ' {

  if not enabled then
     exit sub
  end if

  call openFile()
  print #debugFile, space(indent_) & txt
  call closeFile()

end sub ' }

public sub indent(txt as string) ' {

  if not enabled then
     exit sub
  end if

  text "{ " & txt
  indent_ = indent_ + 2
end sub ' }

public sub dedent() ' {

  if not enabled then
     exit sub
  end if

  indent_ = indent_ - 2
  text "}"

end sub ' }

function enabled() as boolean ' {

' if environ$("username") = "René" and environ$("computername") = "THINKPAD" then
     enabled = true
' else
'    enabled = false
' end if

end function ' }

private sub openFile() ' {

  if not enabled then
     exit sub
  end if

  debugFile = freeFile()
  open fileName_ for append as #debugFile

end sub ' }

private sub closeFile() ' {

  if not enabled then
     exit sub
  end if

  close #debugFile
  debugFile = 0
end sub ' }

