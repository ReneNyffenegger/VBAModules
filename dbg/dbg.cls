'
' vim: ft=vb
'
' dbg.cls
'
option explicit

private indent_     as integer
private dbgWriter_  as dbgWriter

public sub init(dbgWriter__ as dbgWriter) ' {
    indent_      = 0
    set dbgWriter_  = dbgWriter__
end sub ' }

public sub text(txt as string) ' {
    dbgWriter_.out space(indent_) & txt
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

  if indent_ < 0 then
     indent_ = 0
     text "! dbg: Warning, indent_ was set to 0"
  end if
  text "}"

end sub ' }

function enabled() as boolean ' {

' if environ$("username") = "René" and environ$("computername") = "THINKPAD" then
     enabled = true
' else
'    enabled = false
' end if

end function ' }
