option explicit

private debugFile   as integer
private indent      as integer
private fileName_   as string

public sub dbgFileName(fileName as string) ' {
  fileName_ = fileName
end sub ' }

public sub dbg(text as string) ' {

  if not debugEnabled then
     exit sub
  end if

  call startDbg()
  print #debugFile, space(indent) & text
  call endDbg()

end sub ' }

public sub dbgS(text as string) ' {

  if not debugEnabled then
     exit sub
  end if

' indent = indent + 1
  call dbg("{ " & text)
  indent = indent + 2
end sub ' }

public sub dbgE() ' {

  if not debugEnabled then
     exit sub
  end if

  indent = indent - 2
  call dbg("}")
' indent = indent - 1
end sub ' }

function debugEnabled() as boolean ' {

' if environ$("username") = "René" and environ$("computername") = "THINKPAD" then
     debugEnabled = true
' else
'    debugEnabled = false
' end if

end function ' }

private sub startDbg() ' {

  if not debugEnabled then
     exit sub
  end if

  debugFile = freeFile()
' open debugFileName for output as #debugFile
  open fileName_ for append as #debugFile

end sub ' }

private sub endDbg() ' {

  if not debugEnabled then
     exit sub
  end if

  close #debugFile
end sub ' }

