'
'
'   1   Create a module, copy paste content of this file into new module
'
'   2   In «Immediate Window» (german: «Direktfenster») to be found under Menu «View» (german: «Ansicht»)
'       or using Ctrl-G) , first do
'          call application.VBE.activevbProject.references.addFromGuid ("{0002E157-0000-0000-C000-000000000046}", 5, 3)

'   2a  Optionally, you might want to rename the newly inserted module
'      (Still in the «immediate window»:
'          vbe.Activevbproject.VBComponents(1).Name = "00_ModuleLoader"
'       2018-06-09: Apparently, that's not so simple in Excel.
'
'   3   Then load the modules by calling
'          call loadOrReplaceModuleWithFile("fooModule", "c:\path\to\modFoo.bas")
'       for each module
'

option explicit

sub loadOrReplaceModuleWithFile(moduleName as string, pathToFile as string, optional moduleType as long = vbext_ct_StdModule) ' {
 '
 '  3rd argument, moduleType: By default, this sub loads a standard module.
 '                In order to load a class module, use vbext_ct_ClassModule.

 '  dim mdl   as module
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

    call loadModuleFromFile(moduleName, pathToFile, moduleType)

end sub ' }

sub loadModuleFromFile(moduleName as string, pathToFile as string, moduleType as long) ' {

    dim vbComp as vbComponent

  '
  ' Seems to always import a «standard» module:
  '
  '   set vbComp = application.VBE.activeVBProject.vbComponents.import(pathToFile)
  '

    set vbComp = application.VBE.activeVBProject.vbComponents.add(moduleType)
    call vbComp.codeModule.addFromFile(pathToFile)

    vbComp.name  = moduleName

'   Doesn't work, unfortunately
'   vbComp.saved = true

'   Neither does this
'      (Run-time error 29068: Microsoft Access cannot complete this operation
'       You must stop the code and try again)
'   doCmd.save acModule, moduleName

'   This does not generate a runtime error, but
'   does not seem to save the module, either.
'   2018-06-09: It appears to be executable with Access only anyway.
'              (And I have forgotten why it was required).
'      doCmd.close acModule, moduleName, acSaveYes

end sub ' }
