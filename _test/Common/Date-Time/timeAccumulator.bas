option explicit

public declare ptrSafe sub Sleep lib "kernel32" (byVal Milliseconds as longPtr)

sub main() ' {

    dim ta_1, ta_2, ta_3 as timeAccumulator

    set ta_1 = new timeAccumulator
    set ta_2 = new timeAccumulator
    set ta_3 = new timeAccumulator

    dim i         as long
    dim totalTime as double

    ta_1.startAccumulating
         for i = 1 to 10000
             ta_2.startAccumulating
             ta_2.stopAccumulating
         next i

    totalTime = ta_1.stopAccumulating

    debug.print "accumulator 1:    " & ta_1.elapsed_ms
    debug.print "accumulator 2:    " & ta_2.elapsed_ms
    debug.print "totalTime:        " & totalTime

    dim approx_1_second, approx_1000_ms as double
    ta_3.startAccumulating
         sleep 1000
    approx_1_second = ta_3.stopAccumulating / ta_3.freq
    approx_1000_ms  = ta_3.elapsed_ms
    debug.print "approx_1_second:  " & approx_1_second
    debug.print "approx_1000_ms:   " & approx_1000_ms

end sub ' }
