option explicit

function slurpFile(fileName as string) as string ' {

   dim f        as integer

   f = FreeFile()
   open fileName for input as #f
   slurpFile = input(lof(f), #f)
   close f

end function ' }

function fileExists(fileName as string) as boolean ' {

' http://stackoverflow.com/a/28237845/180275
  on error resume next
  fileExists = (GetAttr(fileName) And vbDirectory) <> vbDirectory 

end function ' }
