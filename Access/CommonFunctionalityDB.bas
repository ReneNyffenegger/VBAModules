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
'
'   Compare with executeQuery (below)
'

    on error goto err

'
'   2019-01-04:
'     Apparently, it's not possible to create views via currentDB because
'     currentDB returns a DAO object...
'     See https://stackoverflow.com/a/32772851/180275
'
'   call currentDB().execute(stmt, dbFailOnError)
'
'     However, currentProject.connection returns an ADO connection object
'     with which it apprently is possible to create views:
'
    currentProject.connection.execute stmt

' done:
    exit sub

  err:
    call msgBox("CommonFunctionalityDB - executeSQL" & vbCrLf & err.description & " [" & err.number & "]"& vbCrLf & "stmt = " & stmt)

  ' TODO http://www.lazerwire.com/2011/11/excel-vba-re-throw-errorexception.html
    err.raise err.number, err.source, err.description, err.helpFile, err.helpContext
'   resume done
end sub ' }

sub executeQuery(byVal stmt as string) ' {
'
'   Compare with executeSQL (above)
'

    const qryName = "tq84Query"

    dim qry as dao.queryDef
    set qry = createOrReplaceQuery(qryName, stmt)

    doCmd.openQuery qryName

end sub ' }

sub deleteTable(tableName as string) ' {
    call executeSQL("delete from " & tableName)
end sub ' }

sub dropTableIfExists(tablename as string) ' {
    on error resume next
    executeSQL("drop table " & tablename)
end sub ' }

function createOrReplaceQuery(name as string, stmt as string) as dao.queryDef ' {
'
' 2019-01-10: created from sub createQuery
'

  on error resume next
  set createOrReplaceQuery = currentDB().queryDefs(name)
  on error goto 0

  if not createOrReplaceQuery is nothing then
   '
   ' The following sysCmd checks if the query is open.
   ' Apparently, the doCmd.close (below) does not fail if
   ' the query is not opened.
   '
   ' if sysCmd(acSysCmdGetObjectState, acQuery, name) = acObjStateOpen Then

      '
      ' A queryDef can only be deleted if it is closed.
      '
        doCmd.close acQuery, name
   ' end if

     currentDB().queryDefs.delete(name)
  end if

  set createOrReplaceQuery = currentDB().createQueryDef(name, stmt)

end function ' }

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
