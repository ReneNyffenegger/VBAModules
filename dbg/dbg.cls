'
' vim: ft=vb
'
' dbg.cls
'
option explicit

private indent_     as integer
private dbgWriter_  as dbgWriter

private function removeSpecialCharacters(txt as string) ' {

  removeSpecialCharacters = replace(txt                    , "{"    , "[")
  removeSpecialCharacters = replace(removeSpecialCharacters, "}"    , "]")
  removeSpecialCharacters = replace(removeSpecialCharacters, chr(10), " ")
  removeSpecialCharacters = replace(removeSpecialCharacters, chr(13), " ")

end function ' }

public sub init(dbgWriter__ as dbgWriter) ' {
    indent_      = 0
    set dbgWriter_  = dbgWriter__
end sub ' }

private sub text_(txt as string) ' {
  dbgWriter_.out format(now(), "yyyy-mm-dd hh:MM:ss") & " " & space(indent_) & txt
end sub ' }

public sub text(txt as string) ' {
     text_ removeSpecialCharacters(txt)
end sub ' }

public sub indent(txt as string) ' {

  if not enabled then
     exit sub
  end if

  text_ "{ " & removeSpecialCharacters(txt)

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
  text_ "}"

end sub ' }

public sub unhandledError() ' {

    text ("Caught unhandled error: " & err.description)
    dedent
    err.raise err.number, err.source, err.description

end sub ' }

function enabled() as boolean ' {

' if environ$("username") = "René" and environ$("computername") = "THINKPAD" then
     enabled = true
' else
'    enabled = false
' end if

end function ' }
