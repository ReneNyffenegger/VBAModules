'
'  V.2
'
option explicit

' { WinAPI declarations

declare function GlobalAlloc         lib "kernel32"                                   ( _
     byVal wFlags         as long, _
     byVal dwBytes        as long) as long

declare function lstrcpy lib "kernel32"                                               ( _
     byVal lpString1 as any, _
     byVal lpString2 as any) as long


declare function EmptyClipboard      lib "User32" () as long

declare function CloseClipboard      lib "User32" () as long
declare function OpenClipboard       lib "User32"                                     ( _
     byVal hwnd          as long                   ) as long

declare function GlobalLock          lib "kernel32"                                   ( _
     byVal hMem          as long                   ) as long

declare function GlobalUnlock        lib "kernel32"                                   ( _
     byVal hMem          as long                   ) as long

declare function SetClipboardData    lib "User32"                                     ( _
     byVal wFormat as long, _
     byVal hMem    as long                         ) as long

declare function GetClipboardData    lib "User32"                                     ( _
     byVal wFormat as long                         ) as long

private const GHND                          = &h42
private const CF_TEXT                       = 1

' }

sub textToClipboard(txt as string) ' {

   dim memory       as long
   dim lockedMemory as long

   memory       = GlobalAlloc(GHND, len(txt) + 1)
   if memory = 0 then
      msgBox "GlobalAlloc failed"
      exit sub
   end if

   lockedMemory = GlobalLock(memory)
   if lockedMemory = 0 then
      msgBox "GlobalLock failed"
      exit sub
   end if

   lockedMemory = lstrcpy(lockedMemory, txt)

   call GlobalUnlock(memory)

   if openClipboard(0) = 0 Then
      msgBox "openClipboard failed"
      exit sub
   end if

   EmptyClipboard

   call SetClipboardData(CF_TEXT, memory)

   if CloseClipboard() = 0 then
      msgBox "CloseClipboard failed"
   end if

end sub ' }

function clipboardToText() as string ' {

   dim h            as long
   dim lockedMemory as long


   if OpenClipboard(0) = 0 then
      msgBox "Could not open Clipboard"
      exit function
   end if

   h = GetClipboardData(CF_TEXT)

   if h = 0 then
      msgBox "GetClipboardData could not allocate memory"
      exit function
   end if

   lockedMemory = GlobalLock(h)
   if lockedMemory = 0 then
      msgBox "GlobalLock failed"
      exit function
   end if

   dim str as string

 ' Hopefully, 1 MB is enough...
 ' ... otherwise, the lstrcpy below will fail miserably!
   str = space$(1# * 1024 * 1024)
   
   lstrcpy str, lockedMemory
   GlobalUnlock(lockedMemory)

 '
 ' Remove 'text' after null-byte:
 '
   dim pos as long
   pos = inStr(1, str , Chr$(0), 0)
   str = mid(str, 1,  pos-1)

   if CloseClipboard() = 0 then
      msgBox "CloseClipboard failed"
   end if

   clipboardToText = str

end function ' }
