option explicit

private declare function win32_GetTempPath lib "kernel32" Alias "GetTempPathA" (byVal nBufferLength as long, byVal lpBuffer As string) as long

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

public function tempPath() as string ' {
    const MAX_PATH = 260
    tempPath = string$(MAX_PATH, chr$(0))
    win32_GetTempPath MAX_PATH, tempPath
    tempPath = replace(tempPath, chr$(0), "")
end function ' }
