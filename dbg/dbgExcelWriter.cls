' vi: ft=vb
option explicit

implements dbgWriter

private worksheetOut       as worksheet
private curRow             as long

public sub class_initialize() ' {
       curRow = 0
end sub ' }

public sub class_terminate() ' {
end sub ' }

public sub init(wb as workbook, worksheetName as string) ' {

    set worksheetOut     = wb.worksheets.add
    worksheetOut.visible = xlSheetVeryHidden
    worksheetOut.name    = worksheetName

end sub ' }

public sub dbgWriter_out(txt as string) ' {

    curRow = curRow + 1
    worksheetOut.cells(curRow, 1).value     =  txt
    worksheetOut.cells(curRow, 1).font.name = "Courier New"

end sub ' }
