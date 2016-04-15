sub runSQLScript(pathToScript as string)

  dim sqlText as string

  sqlText = slurpFile(pathToScript)

  sqlText = removeSQLComments(sqlText)

  dim sqlStatements() as string

  sqlStatements = strings.split(sqlText, ";")

  dbgFileName(currentProject.path & "\log\sql")

  dim i as long
  for i = lbound(sqlStatements) to ubound(sqlStatements) - 1 ' Last "statement" is empty because split also returns the part after the last ; -> skip it
      dbg("sqlStatement = " & sqlStatements(i))
      call executeSQL(sqlStatements(i))
  next i


end sub
