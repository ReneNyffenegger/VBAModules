option explicit


' copyExcelSheetToNewAccessTable {       
public sub copyExcelSheetToNewAccessTable( _
  excelFile     as string,                 _
  sheetName     as string,                 _
  accessFile    as string,                 _
  newTableName  as string)

' Compare ..\Access\CommonFunctionalityDB.bas -> importExcelDataIntoTable

  dim con as new ADODB.connection

  con.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFile & ";Extended Properties=Excel 8.0"
  con.open

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

  con.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFileFrom '  & ";Extended Properties=Excel 8.0"
  con.open

  con.execute("select * into [" & tableNameTo & "] in '" & accessFileTo & "' from [" & tableNameFrom & "]") 

end sub ' }
