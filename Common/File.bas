option explicit

function slurpFile(fileName as string) as string ' {

   dim f        as integer

   f = FreeFile()
   open fileName for input as #f
   slurpFile = input(lof(f), #f)
   close f

end function ' }
