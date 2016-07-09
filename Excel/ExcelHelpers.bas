function colLetterToNum(colLetter As String) As Long
' http://vba4excel.blogspot.ch/2012/12/column-number-to-letter-and-reverse.html
  colLetterToNum = activeWorkbook.worksheets(1).columns(colLetter).Column
end function
