'
'      CommonFunctionalityDB
'

' 2019-07-25: It seems that newer versions of Access
'   automatically add «option cmopare database»
'   when a VBA module is inserted programmatically.
'   Therefore: commenting it out:
'
' option compare database
'

option explicit

function getRS(stmt as string) as dao.recordSet ' {
  ' set getRS = dbEngine.workspaces(0).databases(0).openRecordset(stmt)
    set getRS = currentDB().openRecordset(stmt)
end function ' }

function executeSQL(byVal stmt as string) as long ' {
'
'   Returns the numbers of rows affected.
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
    currentProject.connection.execute stmt, executeSQL

' done:
    exit function

  err:
    call msgBox("CommonFunctionalityDB - executeSQL" & vbCrLf & err.description & " [" & err.number & "]"& vbCrLf & "stmt = " & stmt)

  ' TODO http://www.lazerwire.com/2011/11/excel-vba-re-throw-errorexception.html
    err.raise err.number, err.source, err.description, err.helpFile, err.helpContext
'   resume done
end function ' }

sub executeQuery(byVal stmt as string) ' {
'
'   Executes an SQL query and shows the result
'   in a grid.
'
'   Compare with executeSQL (above)
'   

    on error goto err_

    const qryName = "tq84Query"

    dim qry as dao.queryDef
    set qry = createOrReplaceQuery(qryName, stmt)

    doCmd.openQuery qryName

    exit sub
  err_:
    msgBox("Error in executeQuery: " & err.description & " (" & err.source & ")")
    showErrors
end sub ' }

sub executeQueryFromFile(fileName as string) ' {

    executeQuery(removeSQLComments(slurpFile(fileName)))

end sub ' }

sub runSQLScript(pathToScript as string) ' {

    dim sqlStatements() as string

  '
  ' sqlStatementsOfFile() is found in ../Database/SQL.bas ( development/languages/VBA/modules/Database/SQL )
  '
    sqlStatements = sqlStatementsOfFile(pathToScript)

  ' dbgFileName(currentProject.path & "\log\sql")

    dim i as long
    for i = lbound(sqlStatements) to ubound(sqlStatements) - 1 ' Last "statement" is empty because split also returns the part after the last ; -> skip it
     ' dbg("sqlStatement = " & sqlStatements(i))
       call executeSQL(sqlStatements(i))
    next i

end sub ' }

sub deleteTable(tableName as string) ' {
    call executeSQL("delete from " & tableName)
end sub ' }

function doesTableExist(tableName as string) as boolean ' {

    if isNull(dLookup("Name", "MSysObjects", "Name='" & tableName & "'")) then
       doesTableExist = false
    else
       doesTableExist = true
    end if

end function ' }

sub dropTableIfExists(tableName as string) ' {

  if not doesTableExist(tableName) then ' {
     exit sub
  end if ' }

  ' 2019-01-30:  on error resume next

  '
  ' First: close the potentially table in order to prevent error
  '   »The database engine could not lock table '…' because it is already
  '    in use by another person or process [-2147217900]«
  '
    doCmd.close acTable, tableName, acSaveNo

  '
  ' Then: drop table
  '
    executeSQL("drop table " & tableName)
end sub ' }

function truncDate(dt as variant) as variant ' {

    if isNull(dt) then
       truncDate = null
       exit function
    end if

  '
  ' Add the numbers of seconds per day minus one
  ' to dt and round down.
  '
    truncDate = cDate(fix(dateAdd("s", 86399, dt)))
end function ' }

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
        doCmd.close acQuery, name, acSaveNo
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

sub closeAllQueryDefs() ' {

    dim qry as dao.queryDef
    for each qry in currentDb().queryDefs
        doCmd.close acQuery, qry.name, acSaveNo
    next qry

end sub ' }

function nvl2(val as variant, retIfNotNull as variant, retIfNull as variant) ' {
  '
  ' Simulate Oracle's nvl2 function
  '
    if isNull(val) then
       nvl2 = retIfNull
       exit function
    end if

    nvl2 = retIfNotNull

end function ' }

function eq(val_1 as variant, val_2 as variant) as boolean ' {

     if isNull(val_1) then
        if isNull(val_2) then
           eq = true
        else
           eq = false
        end if
        exit function
    end if

    if isNull (val_2) then
       eq = false
       exit function
    end if

    eq = val_1 = val_2 

end function ' }

sub diff2recordSetsRecords(rs1 as dao.recordSet, rs2 as dao.recordSet) ' {

    dim fld as variant: for each fld in rs1.fields ' {

        if isNull(fld.value) and not isNull(rs2(fld.name)) then ' {
           debug.print "diff in " & fld.name & ": " & fld.value & " <> " & rs2(fld.name)
           goto next_fld
        end if ' }

        if not isNull(fld.value) and isNull(rs2(fld.name)) then ' {
           debug.print "diff in " & fld.name & ": " & fld.value & " <> " & rs2(fld.name)
           goto next_fld
        end if ' }

        if typeName(fld.value) = "Double" then ' {
           if round(fld, 6) <> round(rs2(fld.name), 6) then
              debug.print "diff in " & fld.name & ": " & fld.value & " <> " & rs2(fld.name)
           end if ' }
        else ' {
           if fld <> rs2(fld.name) then
              debug.print "diff in " & fld.name & ": " & fld.value & " <> " & rs2(fld.name)
           end if
        end if ' }

next_fld:

    next fld ' }

end sub ' }

sub showErrors() ' {

    dim e as dao.error
    for each e in dbEngine.errors
        debug.print(e.source & ": " & e.description)
    next e

end sub ' }
