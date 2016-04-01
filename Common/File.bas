option explicit

function slurpFile(fileName as string) as string ' {

   dim f        as integer

   f = FreeFile()
   open fileName for input as #f
   slurpFile = input(lof(f), #f)
   close f

end function ' }

function fileBaseName(filename as string) as string ' {

  dim fso as new fileSystemObject

  fileBaseName = fso.getBaseName(filename)

end function ' }

function fileSuffix(filename as string) as string ' {

  dim fso as new fileSystemObject

  fileSuffix = fso.getExtensionName(filename)

end function ' }

function fileExists(fileName as string) as boolean ' {

' http://stackoverflow.com/a/28237845/180275
  on error resume next
  fileExists = (GetAttr(fileName) And vbDirectory) <> vbDirectory 

end function ' }
