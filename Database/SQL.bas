public function removeSQLComments(sqlText as string) as string ' {
'
'   http://stackoverflow.com/a/1839290/180275
'
'   TODO: Currently removes -- only, but should also
'   remove /* ... */

  dim re as new regExp
  set re = createObject("vbscript.regexp")

  re.pattern   = "--.*$"
  re.global    = true
  re.multiLine = true

  removeSQLComments = re.replace(sqlText, "")

end function ' }
