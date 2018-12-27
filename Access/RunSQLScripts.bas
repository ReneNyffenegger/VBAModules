'
'      TODO: This module should probably be merged with CommonFunctionalityDB.bas
'

option explicit

sub runSQLScript(pathToScript as string) ' {

    dim sqlStatements() as string
    sqlStatements = sqlStatementsOfFile(pathToScript)

  ' dbgFileName(currentProject.path & "\log\sql")

    dim i as long
    for i = lbound(sqlStatements) to ubound(sqlStatements) - 1 ' Last "statement" is empty because split also returns the part after the last ; -> skip it
     ' dbg("sqlStatement = " & sqlStatements(i))
       call executeSQL(sqlStatements(i))
    next i

end sub ' }
