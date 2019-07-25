option explicit

public function removeSQLComments(sqlText as string) as string ' {
'
'   http://stackoverflow.com/a/1839290/180275
'
'   Needs Regexp refereence reference:
'       call application.VBE.activeVBProject.references.addFromGuid("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5)
'
'   TODO: Currently removes -- only, but should also
'   remove /* ... */

  dim re as new regExp

  re.pattern   = "--.*$"
  re.global    = true
  re.multiLine = true

  removeSQLComments = re.replace(sqlText, "")

end function ' }

public function sqlStatementsOfFile(pathToScript as string, optional charSet as string = "utf-8") as string() ' {

   dim sqlText as string

 ' Find slurpFileCharSet() @ https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/File
   sqlText = slurpFileCharSet(pathToScript, charSet)
 
   sqlText = removeSQLComments(sqlText)

   dim sqlStatements() as string

 '
 ' This is of course shaky, at best, but up to
 ' now, it did the job:
 '
   sqlStatementsOfFile = strings.split(sqlText, ";")

end function ' }
