'
' vi: ft=basic
'
' Measure accumulated time
'
' https://stackoverflow.com/a/198702/180275 was helpful
'

private type LARGE_INTEGER ' { winnt.h
    LowPart  as long
    HighPart as long
end type ' }

private declare function QueryPerformanceCounter   lib "kernel32" (lpPerformanceCount as LARGE_INTEGER) as long ' profileapi.h
private declare function QueryPerformanceFrequency lib "kernel32" (lpFrequency        as LARGE_INTEGER) as long

const two_32 as double = 4294967296# ' 256^4

private freq            as double
private accumulatedTime as double
private startedTime     as double

private sub class_initialize() ' {
    accumulatedTime = 0

    dim freq_ as LARGE_INTEGER
    QueryPerformanceFrequency freq_

    freq = LARGE_INTEGER_2_double(freq_)

end sub ' }

private function LARGE_INTEGER_2_double(li as LARGE_INTEGER) as double ' {

    LARGE_INTEGER_2_double = li.LowPart
    if LARGE_INTEGER_2_double < 0 then
       LARGE_INTEGER_2_double = LARGE_INTEGER_2_double + two_32
    end if

    LARGE_INTEGER_2_double = li.HighPart * two_32 + LARGE_INTEGER_2_double

end function ' }

public sub startAccumulating() ' {

    dim li as LARGE_INTEGER
    QueryPerformanceCounter li
    startedTime = LARGE_INTEGER_2_double(li)

end sub ' }

public sub stopAccumulating() ' {

    dim li as LARGE_INTEGER
    QueryPerformanceCounter li
    accumulatedTime = accumulatedTime + LARGE_INTEGER_2_double(li) - startedTime

end sub ' }

public function elapsed_ms() as double ' {
    elapsed_ms = 1000# * accumulatedTime / freq
end function ' }
