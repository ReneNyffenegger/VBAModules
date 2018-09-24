'
'  vim: ft=vb
'
option explicit

implements dbgWriter

public sub class_terminate() ' {
  ' Currently not used
end sub ' }

public sub init() ' {
  ' Currently not used
end sub ' }

public sub dbgWriter_out(txt as string) ' {
    selection.typeText txt & chr(10)
end sub ' }
