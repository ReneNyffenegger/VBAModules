option explicit

public function regexpSplit(text as string, pattern as string) as string() ' {

  dim text_0 as string
  dim re     as new regExp

  re.pattern   = pattern
  re.global    = true
  re.multiLine = true

  text_0 = re.replace(text, vbNullChar)

  regexpSplit = strings.split(text_0, vbNullChar)

end function ' }

function lpad(text as String, length as integer, optional padChar as string = " ") ' {
    lpad = string(length - len(text), padChar) & text
end function ' }

function rpad(text as String, length as integer, optional padChar as string = " ") ' {
    rpad = text & string(length - len(text), padChar)
end function ' }
