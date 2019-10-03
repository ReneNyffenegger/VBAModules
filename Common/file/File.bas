option explicit

private declare ptrSafe function win32_GetTempPath lib "kernel32" Alias "GetTempPathA" (byVal nBufferLength as long, byVal lpBuffer As string) as long

function slurpFile(fileName as string) as string ' {
 '
 ' 2019-01-29: it turns out, VBA cannot really cope with
 ' reading utf 8 encoded files.
 ' Therefor, the ADODB substitue slurpFileCharSet might
 ' be considered instead.
 '

   dim f as integer
   f = freeFile()

   open fileName for input as #f

   slurpFile = input(lof(f), #f)

   close f

end function ' }

function slurpFileCharSet(fileName as string, optional charSet as string = "utf-8") as string ' {
  '
  ' Read (slurp) a file in a specific charset.
  '
  ' Needs the ADODB reference:
  '   call application.VBE.activeVBProject.references.addFromGuid("{B691E011-1797-432E-907A-4D8C69339129}", 6, 1)
  '
    dim s as new adodb.stream
    s.charSet = charSet

    s.open
    s.loadFromFile(filename)

    slurpFileCharSet = s.readText

    s.close

end function ' }

sub flushToFile(filename as string, txt as string) ' {

   dim f as integer
   f = freeFile()

   open fileName for output as #f

   print# f, txt

   close f

end sub ' }

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
  fileExists = (getAttr(fileName) and vbDirectory) <> vbDirectory 

end function ' }

public function tempPath() as string ' {
 '
 '  2019-01-11: It might probably be easier to
 '              just use:
 '                 environ$("TEMP")
 '
    const MAX_PATH = 260
    tempPath = string$(MAX_PATH, chr$(0))
    win32_GetTempPath MAX_PATH, tempPath
    tempPath = replace(tempPath, chr$(0), "")
end function ' }
