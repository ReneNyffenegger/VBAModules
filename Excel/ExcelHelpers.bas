option explicit

function findWorksheet(name as string, optional deleteIfExists as boolean = false) as excel.worksheet ' {
 '
 '  TODO: https://renenyffenegger.ch/notes/Microsoft/Office/Excel/Object-Model/Worksheet/index -> getWorksheet.bas
 '
 '  Return worksheet with the given name.
 '  If it doesn't exist, it is created.
 ' 
 '  Optionally, deleteIfExists can be set to true to delete an existing worksheet
 '  of the given name prior to creating it
 ' 
    if deleteIfExists then ' {
       deleteWorksheet name
    end if ' }

    on error goto createWorksheet
       set findWorksheet = thisWorkbook.sheets(name)

    '  No error: Worksheet exists. We can return
       exit function

    createWorksheet:
    '  Error encountered, probably because the worksheet didn't exist.
    '  We have to create the worksheet
       set findWorksheet = thisWorkbook.sheets.add(after := thisWorkbook.sheets(thisWorkbook.sheets.count))
           findWorksheet.name = name

end function ' }

sub deleteWorksheet(name_ as string) ' {
 '
 '  https://stackoverflow.com/a/31475530/180275
 '
    dim i as long : for i = sheets.count to 1 step -1 ' {
     '
     '  We're trying to delete a worksheet… therefore
     '  we loop backward.
     '
        if sheets(i).name = name_ then ' {
            application.displayAlerts = false
            sheets(i).delete
            application.displayAlerts = true
        end if ' }

    next i ' }

end sub ' }

sub freezeHeader(ws as excel.workSheet, optional bottomRow as long = 1) ' {

    ws.activate
    ws.rows(bottomRow + 1).select
    activeWindow.freezePanes = true

end sub ' }

function colLetterToNum(colLetter as string) as long ' {
 '
 '  http://vba4excel.blogspot.ch/2012/12/column-number-to-letter-and-reverse.html
 '
    colLetterToNum = activeWorkbook.worksheets(1).columns(colLetter).column

end function ' }
