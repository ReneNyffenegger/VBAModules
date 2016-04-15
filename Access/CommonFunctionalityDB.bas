'
'      CommonFunctionalityDB
'

option compare database
option explicit

function getRS(stmt as string) as dao.recordSet ' {
  ' set getRS = dbEngine.workspaces(0).databases(0).openRecordset(stmt)
    set getRS = currentDB().openRecordset(stmt)
end function ' }

sub executeSQL(byVal stmt as string) ' {

    on error goto err

'   call dbEngine.workspaces(0).databases(0).execute(stmt, dbFailOnError)
    call currentDB().execute(stmt, dbFailOnError)

' done:
    exit sub

  err:
    call msgBox("CommonFunctionalityDB - executeSQL" & vbCrLf & err.description & " [" & err.number & "]"& vbCrLf & "stmt = " & stmt)

  ' TODO http://www.lazerwire.com/2011/11/excel-vba-re-throw-errorexception.html
    err.raise err.number, err.source, err.description, err.helpFile, err.helpContext
'   resume done
end sub ' }

sub deleteTable(tableName as string) ' {
    call executeSQL("delete from " & tableName)
end sub ' }

sub dropTableIfExists(tablename as string) ' {
    on error resume next
    executeSQL("drop table " & tablename)
end sub ' }

sub createQuery(name as string, stmt as string) ' {

  dim qry as queryDef

  on error resume next
  set qry = currentDB().queryDefs(name)

  if not qry is nothing then
     currentDB().queryDefs.delete(name)
  end if

  set qry = currentDB().createQueryDef(name, stmt)

end sub ' }

function singleSelectValue(stmt as string) as variant ' {

' Return the one row, one column value of
' a select statement, such as in «select count(*) from x»

  dim rs as dao.recordSet
  set rs = getRS(stmt)
  singleSelectValue = rs(0)
  set rs = nothing

end function ' }

sub importExcelDataIntoTable(tablename as string, pathToWorkbook as string, worksheet as string, optional range as string = "", optional hasFieldNames as boolean = false) ' {

  dim worksheet_range as string

  if worksheet = "" then
     worksheet_range = ""
  else
     worksheet_range = worksheet & "!" & range
  end if

' use acLink to link to the data
' doCmd.transferSpreadsheet acImport, , tablename, pathToWorkbook, hasFieldNames, worksheet & "!" & range
  doCmd.transferSpreadsheet acImport, , tablename, pathToWorkbook, hasFieldNames, worksheet_range

end sub ' }

sub importAccessDataIntoTable(tablename as string, pathToDB as string, tablenameSource as string, optional hasFieldNames as boolean = false) ' {

  on error goto nok

' use acLink to link to the data

  doCmd.transferDatabase acImport, "Microsoft Access" , pathToDB, acTable, tablenameSource, tablename

  done:
  exit sub

  nok:
  call err.raise(vbObjectError + 1000, "CommonFunctionalityDB.bas", _ 
     err.description                        & vbCrLf & _
     "err.number = "      & err.number      & vbCrLf & _
     "tableName = "       & tableName       & vbCrLf & _
     "pathToDB  = "       & pathToDB        & vbCrLf & _ 
     "tablenameSource = " & tablenameSource & vbCrLf)

end sub ' }
