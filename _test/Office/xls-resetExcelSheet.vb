option explicit

sub main() ' {
    reset_excel_sheet
end sub ' }

sub reset_excel_sheet() ' {
    resetExcelSheet activeSheet

    createButton                          _
       range(cells( 2, 4), cells( 3, 5)), _
      "Init Sheet",                       _
      "init_sheet"

end sub ' }

sub init_sheet() ' {

    activeWindow.splitRow = 6
    activeWindow.panes(2).scrollRow    = 36
    activeWindow.panes(2).scrollColumn = 16

    createButton                          _
       range(cells(40,19), cells(41,20)), _
      "Reset Excel Sheet",                _
      "reset_excel_sheet"


    range(cells(38, 18), cells(39,21)).interior.color = rgb(250, 180, 30)


end sub ' }
