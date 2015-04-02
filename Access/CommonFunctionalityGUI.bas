'
'    commonFunctionalityGUI
'
option explicit

sub stopFlashingWhileCreatingForm()
    application.VBE.mainWindow.visible = false
end sub

sub openFormDesignHidden(formName as string)
    doCmd.openForm formName, acDesign, , , , acHidden
end sub

function doesFormExist as boolean

    doesFormExist = ("TODO"="TODO")

end function

sub createForm_(name as string)

    dim frm as form
    set frm = createForm

    dim name_orig as string: name_orig = frm.name

    doCmd.save acForm, name_orig
'   call closeForm(frm)
    doCmd.close acForm, name_orig, acSaveYes
    doCmd.rename name, acForm, name_orig

end sub

' sub renameForm(frm as form, newName as string)
'     
'     doCmd.rename newName, acForm, frm.name
' 
' end sub

'sub closeForm(frm as form)
'
'     doCmd.close acForm, frm.name, acSaveYes
'
'end sub

function createLabel(               _ 
            formName as string    , _
            section  as acSection , _
            x        as long      , _
            y        as long      , _
            w        as long      , _
            h        as long      , _
            caption  as string    ) as access.label

 

    set createLabel=createControl(formName, acLabel, section, , , x, y, w, h)
    createLabel.caption = caption

end function

function createTextBox(               _
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long     , _
             controlSrc as string ) as access.textBox

    set createTextBox = createControl(formName, acTextbox, acDetail, , , x, y, w, h)
    createTextBox.controlSource = controlSrc


end function 

function createNavigationControl (               _
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long       _
            ) as access.navigationControl

    set createNavigationControl = createControl(formName, acNavigationControl, acDetail, , , x, y, w, h)

end function

function createNavigationButton (               _
             formName   as string,             _
             navCtl     as navigationControl,  _
             capt       as string _
            ) as access.navigationButton

    set createNavigationButton = createControl(formName, acNavigationButton, acDetail, navCtl.name, , 0, 0, 0, 0)
    createNavigationButton.caption = text

end function

function createTabCtrl(               _
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long       _
            ) as access.tabControl

    set createTabCtrl = createControl(formName, acTabCtl, acDetail, , , x, y, w, h)

end function

sub conditionalFormattingEQStr(tb as textbox, str as string, bgColor as long, fgColor as long)

    dim fc as formatCondition
    set fc = tb.formatConditions.add(acFieldValue, acEqual, """" & str & """")
    fc.backColor = bgColor
    fc.foreColor = fgColor

end sub

sub removeAllControlsOnForm(f as form)

    dim cnt as long

    cnt = f.controls.count

    dim ctrlNo as long

    for ctrlNo = cnt-1 to 0 step -1

        deleteControl f.name, f.controls(ctrlNo).name

    next ctrlNo

end sub

function cm2pt(cm as double) as long

    cm2pt = cm * 567

end function
