option explicit

sub test_moduleLoader() ' {

    dim vbComp as vbComponent
    set vbComp = loadOrReplaceModuleWithFile("loadedModule", activeWorkbook.path & chr(92) & "loaded-module.bas")

    vbComp.codeModule.replaceLine 8, "const num = 42 ' Changed at " & format(now, "yyyy-mm-dd hh:nn")

'   Must be run indirectly in order to prevent compile error
'      «Sub or Function not defined»
    application.run "subInLoadedModule"
'   subInLoadedModule

end sub ' }
