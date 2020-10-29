option explicit

global dbg_ as dbg

sub main(connectionString as variant) ' {

    set dbg_ = new dbg
    dim fileWriter As New dbgFileWriter
    fileWriter.init environ$("TEMP") & "\" & format(now, "yyyy-mm-dd_hhnn") & ".txt", flushImmediately := true
    dbg_.init fileWriter

    dim conn as new adodb.connection

    debug.print("connection string = " & connectionString)
    conn.open connectionString

    createTestTable conn
    insertData      conn
    selectData      conn

    debug.print "finished."

end sub ' }

sub createTestTable(conn as adodb.connection) ' {

    on error resume next
    conn.execute "drop table vba_ado_test"
    on error goto 0

    conn.execute _
    "create table vba_ado_test (                                        " & _
    "  num      number  ( 2) primary key,                               " & _
    "  txt      varchar2(20) not null,                                  " & _
    "  is_prime char(1)      not null    check (is_prime in ('y', 'n'))," & _
    "  dat      date             null                                   " & _
    ")"

end sub ' }

sub insertData(conn as adodb.connection) ' {

    dim insStmt as new adoStatement
    insStmt.init conn

    insStmt.sql "insert into vba_ado_test values(:num, :txt, :is_prime, :dat)"
    insStmt.defineParameters _
       adInteger,     _
       adVarchar, 20, _
       adVarchar,  1, _
       adDate

    conn.beginTrans

    insStmt.exec  1, "one"  , "n", dateSerial(2001,  1,  1)
    insStmt.exec  2, "two"  , "y", dateSerial(2002,  2,  2)
    insStmt.exec  3, "three", "y", dateSerial(2003,  3,  3)
    insStmt.exec  4, "four" , "n", null
    insStmt.exec  5, "five" , "y", null
    insStmt.exec  6, "six"  , "n", dateSerial(2006,  6,  6)
    insStmt.exec  7, "seven", "y", dateSerial(2007,  7,  7)
    insStmt.exec  8, "eight", "n", null
    insStmt.exec  9, "nine" , "n", dateSerial(2009,  9,  9)
    insStmt.exec 10, "ten"  , "n", dateSerial(2010, 10, 10)

    conn.commitTrans

end sub ' }

sub selectData(conn as adodb.connection) ' {

    dim selStmt as new adoStatement
    selStmt.init conn

    selStmt.sql "select * from vba_ado_test where is_prime = :is_prime order by num"

    selStmt.defineParameters adVarchar, 1


    dim rowNum as long

    selStmt.exec "y"
    while selStmt.record ' {

          rowNum = rowNum + 1

          if rowNum = 1 then ' {
             if selStmt("NUM"     ) <> 2                      then msgBox "failure 1.num"
             if selStmt("TXT"     ) <> "two"                  then msgBox "failure 1.txt"
             if selStmt("IS_PRIME") <> "y"                    then msgBox "failure 1.is_prime"
             if selStmt("DAT"     ) <> dateSerial(2002, 2, 2) then msgBox "failure 1.dat"
          end if ' }

          if rowNum = 2 then ' {
             if selStmt("NUM"     ) <> 3                      then msgBox "failure 2.num"
             if selStmt("TXT"     ) <> "three"                then msgBox "failure 2.txt"
             if selStmt("IS_PRIME") <> "y"                    then msgBox "failure 2.is_prime"
             if selStmt("DAT"     ) <> dateSerial(2003, 3, 3) then msgBox "failure 2.dat"
          end if ' }

          if rowNum = 3 then ' {
             if            selStmt("NUM"     ) <> 5           then msgBox "failure 3.num"
             if            selStmt("TXT"     ) <> "five"      then msgBox "failure 3.txt"
             if            selStmt("IS_PRIME") <> "y"         then msgBox "failure 3.is_prime"
             if not isNull(selStmt("DAT"     ))               then msgBox "failure 3.dat"
          end if ' }

          if rowNum = 4 then ' {
             if selStmt("NUM"     ) <> 7                      then msgBox "failure 4.num"
             if selStmt("TXT"     ) <> "seven"                then msgBox "failure 4.txt"
             if selStmt("IS_PRIME") <> "y"                    then msgBox "failure 4.is_prime"
             if selStmt("DAT"     ) <> dateSerial(2007, 7, 7) then msgBox "failure 4.dat"
          end if ' }

          if rowNum > 4 then msgBox "failure: rowNum = " & rowNum
    wend ' }

    selStmt.exec "n"
    while selStmt.record ' {
          rowNum = rowNum + 1

          if rowNum = 5 then ' {
             if selStmt("NUM"     ) <> 1                      then msgBox "failure 5.num"
             if selStmt("TXT"     ) <> "one"                  then msgBox "failure 5.txt"
             if selStmt("IS_PRIME") <> "n"                    then msgBox "failure 5.is_prime"
             if selStmt("DAT"     ) <> dateSerial(2001, 1, 1) then msgBox "failure 5.dat"
          end if ' }

          if rowNum = 6 then ' {
             if            selStmt("NUM"     ) <> 4           then msgBox "failure 7.num"
             if            selStmt("TXT"     ) <> "four"      then msgBox "failure 7.txt"
             if            selStmt("IS_PRIME") <> "n"         then msgBox "failure 7.is_prime"
             if not isNull(selStmt("DAT"     ))               then msgBox "failure 7.dat"
          end if ' }

          if rownum > 10 then msgBox "failure: rowNum = " & rowNum

    wend ' }

    if rowNum <> 10 then msgBox "failure: rowNum = " & rowNum

end sub ' }
