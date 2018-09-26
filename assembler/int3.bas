option explicit

'
'  Execute an INT3 instruction from VBA
'
'  First, a small callback function needs be created, using the init_INT3() function below.
'
'  The starting address of the callback function is stored in the variable callback_INT3.
'
'  The callback function can then be executed with
'    EnumWindows callback_INT3, 0
'
'  The int3 that is caused can then be caught in a debugger such as GDB.
'
'  The WinAPI functions EnumWindows, VirtualAlloc etc. are defined here:
'     https://github.com/ReneNyffenegger/WinAPI-4-VBA/blob/master/WinAPI.bas
'     See also: https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/index

global callback_INT3 as long

sub init_INT3 ' {

    if callback_INT3 = 0 then ' {
       callback_INT3 = VirtualAlloc(0, 9, MEM_RESERVE_AND_COMMIT, PAGE_EXECUTE_RW)
     ' callback_INT3 = HeapAlloc(GetProcessHeap(), 0, 9)

     ' Function's return value
     '
     '     The function's return value is apparently stored
     '     in the EAX register. EnumWindows expects false if
     '     it should not enumerate windows further. Thus,
     '     we load the EAX register with 0 (4 bytes)
     '
       RtlMoveMemory byVal callback_INT3+0, &hB8, 1  ' MOV EAX, …
       RtlMoveMemory byVal callback_INT3+1, &h00, 1  '
       RtlMoveMemory byVal callback_INT3+2, &h00, 1  '
       RtlMoveMemory byVal callback_INT3+3, &h00, 1  '
       RtlMoveMemory byVal callback_INT3+4, &h00, 1  '

     '
     ' The INT 3 instruction
     '
       RtlMoveMemory byVal callback_INT3+5, &hCC, 1  ' INT 3

     '
     ' The return statement that returns from the
     ' callback of EnumWindows.
     '
     ' Since the callback of EnumWindows receives two four
     ' byte parameters (at least in Win32), we additionally
     ' need to pop 8 bytes off the stack:
     '

       RtlMoveMemory byVal callback_INT3+6, &hC2, 1 ' RET (near) with
       RtlMoveMemory byVal callback_INT3+7,    8, 1 '  number of bytes to additionally
       RtlMoveMemory byVal callback_INT3+8,    0, 1 '  pop off the stack

    end if ' }

end sub ' }
