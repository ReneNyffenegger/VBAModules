'
'  Support functions for Excel ranges
'
'  V0.1
'
option explicit

public sub excelRangeAppend(byRef rng as range, rngAdded as range) ' {

    if rng is nothing then
       set rng = rngAdded
       exit sub
    end if

    if not rngAdded is nothing then
       set rng = excel.application.union(rng, rngAdded)
    end if

end sub ' }

' { excelRangeResize
function excelRangeResize (        _
   rng                as range   , _
   optional topRel    as long = 0, _
   optional leftRel   as long = 0, _
   optional bottomRel as long = 0, _
   optional rightRel  as long = 0  _
  ) as range

 '
 ' Function currently assumes that rng is rectangular and consists
 ' of one area.
 '
 ' Test (for the time being) with
 '   excelRangeResize(selection, -1, 1, 1, -3).Interior.Color = rgb(100, 100, 255)

   with rng.parent ' {

       set excelRangeResize = .range (                                                                        _
          .cells(rng.row                  + topRel       , rng.column +                     leftRel      ) , _
          .cells(rng.row + rng.rows.count + bottomRel- 1 , rng.column + rng.columns.count + rightRel  -1 )   _
       )

   end with ' }

end function ' }

function excelRangeExcludeHeader(rng as range) as range ' {
    set excelRangeExcludeHeader = excelRangeResize(rng, topRel := 1)
end function ' }
