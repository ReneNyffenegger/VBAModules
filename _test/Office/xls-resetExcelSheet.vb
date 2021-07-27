option explicit

sub main() ' {
    reset_excel_sheet
end sub ' }

sub reset_excel_sheet() ' {
    resetExcelSheet activeSheet
    createButton range(cells(2,2), cells(3,4)), "Reset Excel Sheet", "reset_excel_sheet"
end sub ' }
