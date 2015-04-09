'
'    commonFunctionalityGUI
'
option explicit

sub stopFlashingWhileCreatingForm() ' {
    application.VBE.mainWindow.visible = false
end sub ' }

sub openFormDesignHidden(formName as string) ' {
    doCmd.openForm formName, acDesign, , , , acHidden
end sub ' }

function doesFormExist(name as string) as boolean ' {

    dim i as long

    for i = 0 to currentDB().containers("Forms").documents.count - 1

        if currentDB().containers("Forms").documents(i).name = name then
           doesFormExist = true
           exit function
        end if

    next i

    doesFormExist = false
end function ' }

function isFormOpen(name as string) as boolean ' {

    dim r as long

    r = sysCmd(acSysCmdGetObjectState, acForm, name)

  ' r is one of:
  '                     0 The object is closed
  '    acObjStateOpen   1 The object is open
  '    acObjStateDirty  2 A change has been made, but unsaved
  '    acObjStateNew    4 The object is new

    if r = 0 then
       isFormOpen = false
       exit function
    end if

    isFormOpen = true

end function ' }

sub createForm_(name as string) ' {

    dim frm as form
    set frm = createForm

    dim name_orig as string: name_orig = frm.name

    doCmd.save acForm, name_orig
'   call closeForm(frm)
    doCmd.close acForm, name_orig, acSaveYes
    doCmd.rename name, acForm, name_orig

end sub ' }

' sub renameForm(frm as form, newName as string) ' {
'     
'     doCmd.rename newName, acForm, frm.name
' 
' end sub ' }

sub deleteForm(name as string) ' {

    if not doesFormExist(name) then
       exit sub
    end if

    if isFormOpen(name) then
       call closeForm(name)
    end if

    doCmd.deleteObject acForm, name

end sub ' }

sub closeForm(name as string) ' {

     doCmd.close acForm, name, acSaveYes

end sub ' }

sub makeContinuous(frm as form) ' {
    frm.defaultView = 1
end sub ' }

 ' createLabel {
function createLabel (              _                 
            formName as string    , _
            section  as acSection , _
            x        as long      , _
            y        as long      , _
            w        as long      , _
            h        as long      , _
            caption  as string    ) as access.label

 

    set createLabel = createControl(formName, acLabel, section, , , x, y, w, h)
    createLabel.caption = caption

end function ' } ' }

 ' createTextBox {
function createTextBox (              _ 
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long     , _
             controlSrc as string ) as access.textBox

    set createTextBox = createControl(formName, acTextbox, acDetail, , , x, y, w, h)
    createTextBox.controlSource = controlSrc


end function ' }

 ' createNavigationControl {
function createNavigationControl (    _
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long       _
            ) as access.navigationControl

    set createNavigationControl = createControl(formName, acNavigationControl, acDetail, , , x, y, w, h)

end function ' }

 ' createNavigationButton {
function createNavigationButton (              _
             formName   as string,             _
             navCtl     as navigationControl,  _
             capt       as string              _
            ) as access.navigationButton

    set createNavigationButton = createControl(formName, acNavigationButton, acDetail, navCtl.name, , 0, 0, 0, 0)
    createNavigationButton.caption = capt

end function ' }

' createTabCtrl {
function createTabCtrl (              _
             formName   as string   , _
             x          as long     , _
             y          as long     , _
             w          as long     , _
             h          as long       _
            ) as access.tabControl

    set createTabCtrl = createControl(formName, acTabCtl, acDetail, , , x, y, w, h)

end function ' }

sub conditionalFormattingEQStr(tb as textbox, str as string, bgColor as long, fgColor as long) ' {

    dim fc as formatCondition
    set fc = tb.formatConditions.add(acFieldValue, acEqual, """" & str & """")
    fc.backColor = bgColor
    fc.foreColor = fgColor

end sub ' }

sub removeAllControlsOnForm(f as form) ' {

    dim cnt as long

    cnt = f.controls.count

    dim ctrlNo as long

    for ctrlNo = cnt-1 to 0 step -1

        deleteControl f.name, f.controls(ctrlNo).name

    next ctrlNo

end sub ' }

function cm2pt(cm as double) as long ' {

    cm2pt = cm * 567

end function ' }

