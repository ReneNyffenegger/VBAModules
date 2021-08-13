'
'  Depends on ../Common/Collection.vb
'
'  V0.10
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
     '
     ' 2021-06-04: it seems safer to use thisWorkbook rather than activeWorkbook
     '
       set wb = thisWorkbook
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

sub deleteWorksheet(name_ as string, optional wb as workbook = nothing)  ' {

    if wb is nothing then
       set wb = thisWorkbook
    end if

    dim ws as worksheet
    set ws = collObjectOrNothing(wb.sheets, name_)

    if not ws is nothing then ' {

     '
     ' Set displayAlerts temporarily to false so that the unwanted message
     '    Microsoft Excel will permanentely delete this sheet. Do you want to continue?
     ' does not pop up.
     '
     ' Compare with another solution on  https://stackoverflow.com/a/31475530/180275
     '

     ' 2021-07-19: Trying to delete very(?) hidden sheets seems not possible
     ' unless sheet is made visible:
       ws.visible = xlSheetVisible

       dim da as boolean : da = application.displayAlerts
       application.displayAlerts = false
       ws.delete
       application.displayAlerts = da

    end if ' }

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

sub freezeHeader(ws as excel.workSheet, optional bottomRow as long = 1, optional leftColumn as long = 0) ' {
  '
  ' TODO: https://stackoverflow.com/a/19362973/180275 seems to indicate that
  ' this sub should make sure that screenUpdating is set to true when the sheet
  ' is frozen
  '
  ' 2021-07-01: Make sure the currently active sheet and range is activated again when the sub
  ' is left
  '
    dim curSheet     as worksheet : set curSheet     = activeSheet

    ws.activate
    dim curSelection as range     : set curSelection = selection

    if leftColumn = 0 then
       ws.rows(bottomRow + 1).select
    else
       ws.cells(bottomRow+1, leftColumn+1).select
    end if

    with activeWindow
         if .freezePanes then .freezePanes = false
'       .splitColumn = 0
'       .splitRow    = bottomRow
        .freezePanes = true
    end with

    curSelection.select

    curSheet.activate

end sub ' }

sub insertHyperlinkToVBAMacro(where as range, byVal text as string, byVal macroname as string, paramArray args()) ' {

    dim formula as string
    formula = "=hyperlink("

    formula = formula & """#" & macroname & "("

    dim firstArgument as boolean : firstArgument = true

    dim argNo as long
    for argNo = lBound(args) to uBound(args) ' {

        if firstArgument then
           firstArgument = false
        else
           formula = formula & application.international(xlListSeparator) ' semicolon or comma
        end if

        if varType(args(argNo)) = vbString then
           formula = formula & """""" & args(argNo) & """"""
        else
           formula = formula & args(argNo)
        end if

    next argNo ' }

    formula = formula & ")"""

    formula = formula & "," ' Always comma, no need to invoke application.international(xlListSeparator)
    formula = formula & """" & text & """"

    formula = formula & ")"

    where.formula = formula
end sub ' }

function colLetterToNum(colLetter as string) as long ' {
 '
 '  http://vba4excel.blogspot.ch/2012/12/column-number-to-letter-and-reverse.html
 '
    colLetterToNum = activeWorkbook.worksheets(1).columns(colLetter).column

end function ' }

function colNumToLetter(colNum as long) as string
 '
 '  http://vba4excel.blogspot.ch/2012/12/column-number-to-letter-and-reverse.html
 '
     colNumToLetter = split(cells(1, colNum).address, "$")(1)
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

function isRibbonShown() as boolean ' {
    isRibbonShown = application.commandBars("Ribbon").controls(1).height >= 100
end function ' }

sub showRibbon(visible as boolean) ' {
 '
 ' Note: Hiding the Ribbon in Excel or Word causes the workbook
 '       or Document to occupy the entire screen.
 '       Thus, before hiding the Ribbon, the size of
 '       application.window (.left, .top etc) might be stored and
 '       applied when the ribbon is shown again.
 '
 '

#if 0 then
 '
 ' This function was originally intended to be put into a
 ' general OfficeHelper VBa-module. However, it turned out
 ' that the differences among Office products are too big
 ' for such a general approach. Thus, this portion of the
 ' excluded with a #if 0 then preprocessor block.
 '
   if application.name = "Microsoft Visio" then
    '
    ' Visio does not seem to have .executeMso "HideRibbon" capability, so it
    ' does not make sense to continue here.
    '
   end if

   if application.name = "Microsoft Access" then ' {
       if visible then doCmd.showToolbar "Ribbon", acToolbarYes _
       else            doCmd.showToolbar "Ribbon", acToolbarNo
       exit sub
   end if ' }

#end if

   if isRibbonShown = visible then
    '
    ' Ribbon already shown/hidden, nothing to be done
    '
      exit sub
   end if

  '
  ' Toggle Ribbon when shown
  '
   dim fs as boolean
   fs = application.displayFullScreen
   application.commandBars.executeMso "HideRibbon"
   application.displayFullScreen = fs

end sub ' }

sub resetExcelSheet(sh as worksheet) ' {

  '
  ' TODO: This function should probably make sure that sh is not protected when called
  '

    sh.columns.useStandardWidth = true
    sh.rows.useStandardHeight   = true
#if 1 then
 '
 '  Drawing a border apparently does not extend
 '  the size of usedRange. Thus, all cells
 '  are cleared and the previous sh.usedRange.clear
 '  left as a reminder
 '
    sh.cells.clear
#else
    sh.usedRange.clear
#end if

    dim shp as shape
    for each shp in sh.shapes ' {
        shp.delete
    next shp ' }

    dim n as name
    for each n in sh.names ' {
        n.delete
    next n ' }

    sh.scrollArea = ""

  '
  ' It seems that a hidden sheet cannot be moved by selecting
  ' a cell on it (well, it sort of makes sense, though).
  '
    dim curSheet as worksheet
    set curSheet = activeSheet
    curSheet.visible = xlSheetVisible
    sh.activate
    sh.cells(1,1).select

    activeWindow.splitRow     = 0
    activeWindow.splitColumn  = 0
    activeWindow.split        = false
    activeWindow.freezePanes  = false

    curSheet.activate

end sub ' }
