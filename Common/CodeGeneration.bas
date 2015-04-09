'
'     CodeGeneration
'
'     Be sure to add "Microsoft Visual Basic for Applications Extensibility 5.3" under References
'

option explicit

sub emptyModuleCodeForForm(frm as form) ' {
    dim mdl as module
    set mdl = frm.module

    dim vbModule as vbComponent
    set vbModule = vbe.activeVBProject.vbComponents.item("Form_" & frm.name)

    dim nofLines as long
    nofLines = vbModule.codeModule.countOfLines

    call vbModule.codeModule.deleteLines(3, nofLines-2)
end sub ' }

sub dynamicEventHandler(frm as form, ctrl as control, codeLine as string, eventName as string) ' {

    dim mdl as module
    set mdl = frm.module

    dim pos as long
    pos = mdl.countOfLines

    call mdl.insertLines(pos+1, "")
    call mdl.insertLines(pos+2, "sub " & ctrl.name & "_" & eventName)
    call mdl.insertLines(pos+3, codeLine)
    call mdl.insertLines(pos+4, "end sub")    

end sub ' }

sub dynamicOnClick(frm as form, ctrl as control, codeLine as string) ' {

    call dynamicEventHandler(frm, ctrl, codeLine, "Click")

end sub ' }

sub replaceModuleWithFile(moduleName as string, pathToFile as string) ' {

    dim mdl as module
    set mdl = VBE.activeVBProject.vbComponents(moduleName)

    if mdl not nothing then
       call VBE.activeVBProject.vbComponents.Remove (VBE.activeProject.vbComponents(moduleName))
    end if

    call loadMOduleFromFile(moduleName, pathToFile)

end sub ' }

sub loadMOduleFromFile(moduleName as string, pathToFile as string) ' {

    dim vbComp as vbComponent
    set vbComp = VBE.activeVBProject.vbComponents.import(pathToFile)

    vbComp.name = moduleName

end sub ' }
