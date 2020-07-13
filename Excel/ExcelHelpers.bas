'
'      Depends on ../Common/Collection.bas
'
option explicit

function findWorksheet(name as string, optional deleteIfExists as boolean = false, optional wb as workbook = nothing) as excel.worksheet ' {
 '
 '  TODO: https://renenyffenegger.ch/notes/Microsoft/Office/Excel/Object-Model/Worksheet/index -> getWorksheet.bas
 '
 '  Return worksheet with the given name.
 '  If it doesn't exist, it is created.
 ' 
 '  Optionally, deleteIfExists can be set to true to delete an existing worksheet
 '  of the given name prior to creating it
 ' 
   
    if wb is nothing then
       set wb = activeWorkbook
    end if

    if deleteIfExists then ' {
       deleteWorksheet name, wb
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

sub deleteWorksheet(name_ as string, wb as workbook) ' {

    dim ws as worksheet
    set ws = collObjectOrNothing(wb.sheets, name_)

    if not ws is nothing then ' {

       dim da as boolean : da = application.displayAlerts

       application.displayAlerts = false

       ws.delete

       application.displayAlerts = da

    end if ' }

'Q '
'Q '  https://stackoverflow.com/a/31475530/180275
'Q '
'Q    dim i as long : for i = sheets.count to 1 step -1 ' {
'Q     '
'Q     '  We're trying to delete a worksheet… therefore
'Q     '  we loop backward.
'Q     '
'Q        if sheets(i).name = name_ then ' {
'Q            application.displayAlerts = false
'Q            sheets(i).delete
'Q            application.displayAlerts = true
'Q        end if ' }
'Q
'Q    next i ' }

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
