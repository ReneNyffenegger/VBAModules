'
'      CommonFunctionalityDB
'

option compare database
option explicit

function getRS(stmt as string) as dao.recordSet ' {
  ' set getRS = dbEngine.workspaces(0).databases(0).openRecordset(stmt)
    set getRS = currentDB().openRecordset(stmt)
end function ' }

sub executeSQL(stmt as string) ' {
'   call dbEngine.workspaces(0).databases(0).execute(stmt, dbFailOnError)
    call currentDB().execute(stmt, dbFailOnError)
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

' use acLink to link to the data
  doCmd.transferSpreadsheet acImport, , tablename, pathToWorkbook, hasFieldNames, worksheet & "!" & range

end sub ' }
