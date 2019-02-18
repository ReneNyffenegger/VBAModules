'
'  Needs the »Microsoft Visual Basic for Applications Extensibility 5.3« reference (for vbProject etc)
'
'    thisWorkbook.VBProject.references.addFromGuid  GUID   :="{0002E157-0000-0000-C000-000000000046}",  major  :=  5,  minor  :=  3
'
option explicit

sub findTextInCode(text as string, optional wholeWord as boolean = false, optional matchCase as boolean = false, optional patternSearch as boolean = false) ' {

 '
 '  Note: matchCase and patternSearch cannot both be true.
 '

    dim proj as vbIDE.vbProject
    for each proj in application.vbe.vbProjects

        if proj.protection <> vbext_pp_locked then ' {
           dim comp as vbIDE.vbComponent
           for each comp in proj.vbComponents

               dim mdl as vbIDE.codeModule

               set mdl = comp.codeModule

               dim startLine as long
               dim startCol  as long
               dim endLine   as long
               dim endCol    as long

            '
            '  The following four values are modified by the .find() procedure below (they're passed "byRef")
            '
               startLine =  1
               startCol  =  1
               endLine   =  mdl.countOfLines
               endCol    = -1

               do while mdl.find( text, startLine, startCol, endLine, endCol, wholeWord, matchCase, patternSearch ) ' {

                  dim projFilename as string
'                 on error resume next
                  projFilename = proj.filename
'                 on error goto 0

'                 if proj.type <> vbext_pt_hostProject then
'                    projFilename = proj.filename
'                 else
'                    projFilename = "..."
'                 end if

                  debug.print("found a match at " & projFileName & " - " & comp.name & " - " & startLine & ":" & startCol)

                  startLine =  endLine
                  startCol  =  endCol
                  endLine   =  mdl.countOfLines
                  endCol    = -1

               loop ' }

           next comp
        end if ' }

    next proj

end sub ' }
