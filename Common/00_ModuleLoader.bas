'
'
'   1   Create a module, copy paste content of this file into new module
'
'   2   In «Immediate Window», first do
'          call Application.VBE.activevbProject.References.AddFromGuid ("{0002E157-0000-0000-C000-000000000046}", 5, 3)
'
'   3   Then load the modules by calling
'          call loadOrReplaceModuleWithFile("fooModule", "c:\path\to\modFoo.bas")
'       for each module
'
sub loadOrReplaceModuleWithFile(moduleName as string, pathToFile as string) ' {

    dim mdl   as module
    dim vbc   as vbComponents
    dim i     as long
    dim found as boolean

    set vbc = application.VBE.activeVBProject.vbComponents

    found = false

    for i = 1 to vbc.count
        if  vbc(i).name = moduleName then
            found = true
            exit for
        end if
    next i

    if found then
       call vbc.remove(vbc(i))
    end if

    call loadMOduleFromFile(moduleName, pathToFile)

end sub ' }

sub loadMOduleFromFile(moduleName as string, pathToFile as string) ' {

    dim vbComp as vbComponent
    set vbComp = application.VBE.activeVBProject.vbComponents.import(pathToFile)

    vbComp.name  = moduleName
'   vbComp.saved = true  ' Doesn't work, unfortunately

end sub ' }
