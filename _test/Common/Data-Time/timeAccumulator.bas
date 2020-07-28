option explicit

sub main() ' {

    dim ta_1, ta_2 as timeAccumulator

    set ta_1 = new timeAccumulator
    set ta_2 = new timeAccumulator

    dim i as long

    ta_1.startAccumulating
         for i = 1 to 10000
             ta_2.startAccumulating
             ta_2.stopAccumulating
         next i
    ta_1.stopAccumulating

    debug.print "accumulator 1: " & ta_1.elapsed_ms
    debug.print "accumulator 2: " & ta_2.elapsed_ms

end sub ' }
