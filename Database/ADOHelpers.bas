option explicit


' copyExcelSheetToNewAccessTable {       
public sub copyExcelSheetToNewAccessTable( _
  excelFile     as string,                 _
  sheetName     as string,                 _
  accessFile    as string,                 _
  newTableName  as string)

' Compare ..\Access\CommonFunctionalityDB.bas -> importExcelDataIntoTable

  dim con as ADODB.connection

  set con = openADOConnectionToExcelFile(excelFile, false)

  con.execute("select * into [" & newTableName & "] in '" & accessFile & "' from [" & sheetName & "$]") 

end sub ' }

' copyAccessTableToNewAccessTable {       
public sub copyAccessTableToNewAccessTable ( _
  accessFileFrom     as string,              _
  tableNameFrom      as string,              _
  accessFileTo       as string,              _
  tableNameTo        as string)

' Compare ..\Access\CommonFunctionalityDB.bas -> importExcelDataIntoTable

  dim con as new ADODB.connection
  set con = openADOConnectionToAccess(accessFileFrom)

  con.execute("select * into [" & tableNameTo & "] in '" & accessFileTo & "' from [" & tableNameFrom & "]") 

end sub ' }

public function openADOConnectionToAccess(accessFile as string) as ADODB.connection ' {

  set openADOConnectionToAccess = new ADODB.connection
  openADOConnectionToAccess.open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & accessFile & "'"

end function ' }

public function openADOConnectionToExcelFile(excelFile as string, optional header as boolean = true) as ADODB.connection ' {

  set openADOConnectionToExcelFile = new ADODB.connection

  openADOConnectionToExcelFile.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & excelFile & "'"

  if header then
     openADOConnectionToExcelFile.connectionString = openADOConnectionToExcelFile.connectionString & ";Extended Properties=""Excel 8.0;HDR=yes;IMEX=1;"""
  else
     openADOConnectionToExcelFile.connectionString = openADOConnectionToExcelFile.connectionString & ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1;"""
  end if

  openADOConnectionToExcelFile.open

end function ' }

' ADODoesTableExist {
public function ADODoesTableExist( _
    con       as ADODB.connection, _
    tableName as string            _
  ) as boolean

' http://stackoverflow.com/a/1082482/180275

  dim rsSchema As ADODB.recordset

  Set rsSchema = con.openSchema(adSchemaColumns, array(empty, empty, tableName, empty))

  if rsSchema.BOF And rsSchema.eof then
     ADODoesTableExist = false
  else
     ADODoesTableExist = true
  end if

  rsSchema.Close
  set rsSchema = Nothing

end function ' }

public function ADOExecuteSQL(con as ADODB.connection, stmt as string) as long ' {
  on error goto nok

    call con.execute(stmt, ADOExecuteSQL)

    exit function

  nok:

    dbgE
    call err.raise(1000 + vbObjectError, "ADOHelper.bas - ADOExecuteSQL",  err.description & " [" & err.number & "]"& vbCrLf & "stmt = " & stmt)

end function ' }

public function ADOSelect1R1C(con as ADODB.connection, stmt as string) as variant ' {

  dim rs as ADODB.recordset

  set rs = con.execute(stmt)

' if not rs.eof then
'    rs.moveNext
     ADOSelect1R1C = rs.fields(0).value
' end if

  rs.close
  

end function ' }
