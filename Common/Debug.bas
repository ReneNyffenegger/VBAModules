option explicit

private debugFile   as integer
private indent      as integer

public sub startDbg(debugFileName as string) ' {

  if not debugEnabled then
     exit sub
  end if

  debugFile = freeFile()
' open debugFileName for output as #debugFile
  open debugFileName for append as #debugFile

end sub ' }

public sub dbg(text as string) ' {

  if not debugEnabled then
     exit sub
  end if

  call startDbg("c:\temp\dbg.txt")
  write #debugFile, space(indent) & text
  call endDbg()

end sub ' }

public sub endDbg() ' {

  if not debugEnabled then
     exit sub
  end if

  close #debugFile
end sub ' }

public sub dbgS(text as string) ' {

  if not debugEnabled then
     exit sub
  end if

  indent = indent + 1
  call dbg("{ " & text)
  indent = indent + 1
end sub ' }

public sub dbgE() ' {

  if not debugEnabled then
     exit sub
  end if

  indent = indent - 1
  call dbg("}")
  indent = indent - 1
end sub ' }

function debugEnabled() as boolean ' {

' if environ$("username") = "René" and environ$("computername") = "THINKPAD" then
     debugEnabled = true
' else
'    debugEnabled = false
' end if

end function ' }
