'
'  V.2
'
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
       range(cells(40,19), cells(41,21)), _
      "Reset Excel Sheet",                _
      "reset_excel_sheet"

    createButton                          _
       range(cells(43,19), cells(44,21)), _
      "insert data validation",           _
      "insert_data_validation"


    range(cells(38, 18), cells(39,21)).interior.color = rgb(250, 180, 30)


end sub ' }

sub insert_data_validation() ' {

    dim rng as range
    set rng = range(cells(43, 23), cells(44, 26))

    dim firstCellRelativeAddress as string
    dim formula                  as string

    firstCellRelativeAddress =  rng.address(rowAbsolute := false, columnAbsolute := false)
    formula                  = "=isNumber(" & firstCellRelativeAddress & ")"

    with rng.validation ' {

        .add type := xlValidateCustom, formula1 := formula

        .ignoreBlank  =  true

        .showInput    =  true
        .inputTitle   = "Validation rule"
        .inputMessage = "Enter a numerical value"

        .showError    =  true
        .errorTitle   = "Validation rule failed"
        .errorMessage = "Please enter a number""

    end with ' }

    rng.borderAround xlContinuous, xlMedium, color := rgb(140, 90, 180)
    rng.cells(1,1).offset(-1) = "Validated data (only numbers) in box below"

end sub ' }
