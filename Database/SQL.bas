option explicit

public function removeSQLComments(sqlText as string) as string ' {
'
'   http://stackoverflow.com/a/1839290/180275
'
'   TODO: Currently removes -- only, but should also
'   remove /* ... */

  dim re as new regExp
' set re = createObject("vbscript.regexp")

  re.pattern   = "--.*$"
  re.global    = true
  re.multiLine = true

  removeSQLComments = re.replace(sqlText, "")

end function ' }

public function sqlStatementsOfFile(pathToScript as string) as string() ' {

   dim sqlText as string

 ' Find slurpFile() @ https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/File
   sqlText = slurpFile(pathToScript)
 
 '
 ' Find removeSQLComments() @ https://renenyffenegger.ch/notes/development/languages/VBA/modules/Database/SQL
   sqlText = removeSQLComments(sqlText)

   dim sqlStatements() as string

 '
 ' This is of course shaky, at best, but up to
 ' now, it did the job:
 '
   sqlStatementsOfFile = strings.split(sqlText, ";")

end function ' }
