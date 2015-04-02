'
'   http://stackoverflow.com/a/1839290/180275
'
sub runSQLScript(pathToScript as string)

   dim vSql       as variant
   dim vSqls      as variant
   dim strSql     as string
   dim intF       as integer

   intF = FreeFile()

   open pathToScript for input As #intF

   dim oRegExp as object
   set oRegExp = createObject("vbscript.regexp")

   oRegExp.pattern   = "--.*$"
   oRegExp.global    = true
   oRegExp.multiLine = true

   dim stmt as string
   stmt = ""
   do until eof(intF)

      dim lin as string
      line input #intF, lin
      lin = oRegExp.replace(lin, "")

      if right$(lin, 1) = ";" then

        lin = left$(lin, len(lin)-1)
        stmt = stmt & lin

      ' executeSQL: see CommonFunctionalityDB.bas
        call executeSQL(stmt)
        stmt = ""

      else

        stmt = stmt & lin

      end if

   loop
  

end sub
