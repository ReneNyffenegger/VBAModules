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

sub deleteRange(name_ as string, optional ws as worksheet = nothing) ' {

    if ws is nothing then
       set ws = activeWorkbook.activeSheet
    end if

 on error goto err_
    dim rng as range
    set rng = ws.range(name_)
 on error goto 0

    rng.clearFormats
    rng.clearContents

    ws.parent.names(name_).delete

    exit sub

 err_:
    if err.number <> 1004 then ' 1004 = Application-defined or object-defined error
        msgBox "deleteRange: " & err.number & chr(10) & err.description
    end if

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

function createButton(rng as range, txt as string, nameSub as string) as button ' {

    set createButton      = rng.parent.buttons.add( left := rng.left, top := rng.top, width := rng.width, height := rng.height)
    createButton.caption  = txt
    createButton.onAction = nameSub

end function ' }

function unprotect(byVal sh as worksheet, byVal pw as string) as boolean ' {

    on error resume next

    sh.unprotect pw

    if err.number = 1004 then ' {
    '
    '  Sheet could not be unprotected
    '
       unprotect = false
       exit function

    end if ' }

    unprotect = true

end function ' }

function pageNumberOfCell(c as range) as long ' {

   dim vPageCnt as integer
   dim hPageCnt as integer

   dim sh as worksheet
   set sh = c.parent

   if sh.pageSetup.Order = xlDownThenOver then
      hPageCnt = sh.hPageBreaks.Count + 1
      vPageCnt = 1
   else

      vPageCnt = sh.vPageBreaks.Count + 1
      hPageCnt = 1

   end if

   pageNumberOfCell = 1

   dim vpb as vPageBreak
   for each vpb In sh.vPageBreaks

       if vpb.Location.Column > c.column then exit for
       pageNumberOfCell = pageNumberOfCell + hPageCnt
   next vpb

   dim hpb as hPageBreak
   for each hpb In sh.hPageBreaks
       If hpb.Location.row > c.row then exit for
       pageNumberOfCell = pageNumberOfCell + vPageCnt
   next hpb

end function ' }
