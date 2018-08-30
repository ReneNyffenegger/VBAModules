option explicit

function colLetterToNum(colLetter as string) as long ' {
  ' http://vba4excel.blogspot.ch/2012/12/column-number-to-letter-and-reverse.html
    colLetterToNum = activeWorkbook.worksheets(1).columns(colLetter).column
end function ' }

function findWorksheet(name as string) as excel.worksheet ' {
  '
  ' Find the worksheet with the given name.
  '
  ' If no such worksheet exists, creates it.
  '

  on error goto noWorksheetFound

    set findWorksheet = thisWorkbook.sheets(name)
    exit function

  noWorksheetFound:
    set findWorksheet = thisWorkbook.sheets.add(after := thisWorkbook.sheets(thisWorkbook.sheets.count))
    findWorksheet.name = name

end function ' }
