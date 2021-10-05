'
'  Support functions for Excel ranges
'
'  V0.4
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

public function excelRangeSubtract(rng as range, rngSub as range) as range ' {
 '
 ' After an idea that I found in https://stackoverflow.com/a/21589364/180275
 '

   if rng is nothing then '
      exit function
   end if

   if rngSub is nothing then
      set excelRangeSubtract = rng
      exit function
   end if

   dim rngCommon  As Range

   set rngCommon = intersect(rng, rngSub)

   if     rngCommon is nothing then
        ' No overlap
          set excelRangeSubtract = rng

   elseIf rngCommon.address = rng.address then
        ' Total overlap
          set excelRangeSubtract = nothing

   else
    '
    '  We have a partial overlap between the
    '  the two ranges.
    '  So, we iterate over each area of rng
    '  and subtract the common area from it.
    '  The result of each subtraction is accumulated
    '  in rngACcumulator.
    '
       dim rngACcumulator as range

       dim rngArea   as range
       for each rngArea in rng.areas

          if intersect(rngArea, rngCommon) is nothing then

             excelRangeAppend rngACcumulator, rngArea

          else

             if rngArea.cells.count = 1 then

              ' Nothing to do?

             else

                dim rngPart_1  as range
                dim rngPart_2  as range

                if rngArea.rows.count > 1 then

                 ' Split the range into a top and bottom half:

                   set rngPart_1 = rngArea.resize(rngArea.rows.count \ 2)
                   set rngPart_2 = rngArea.resize(rngArea.rows.count - rngPart_1.rows.count).offset(rngPart_1.Rows.Count)

                 else

                 ' Split the range into a left and right half:

                   set rngPart_1 = rngArea.resize(, rngArea.columns.count \ 2                   )
                   set rngPart_2 = rngArea.resize(, rngArea.columns.count - rngPart_1.columns.count).offset(, rngPart_1.columns.count)

                 end if

                 excelRangeAppend rngACcumulator, excelRangeSubtract( rngPart_1, rngCommon )
                 excelRangeAppend rngACcumulator, excelRangeSubtract( rngPart_2, rngCommon )

              end if
          end if

       next rngArea

       set excelRangeSubtract = rngACcumulator

   end if

end function ' }

public function excelRangeToJson(rng as range) as string ' {

    dim ret as new stringBuffer
    ret.init 10000

    ret.append "["

    dim r as long
    for r = 1 to rng.rows.count

        if r > 1 then
           ret.append ","
        end if
        ret.append("[")

        dim c as long
        for c = 1 to rng.columns.count

            if c > 1 then
               ret.append ","
            end if

            ret.append json_val(rng.cells(r,c).value)

        next c
        ret.append "]"

    next r

    ret.append "]"

    excelRangeToJson = ret

end function ' }

public sub excelRangeToolTip(rng as range, title as string, msg as string) ' {

   with rng.validation ' {

       .add  type   := xlValidateCustom, formula1 := "=true"
       .inputTitle   = title
       .inputMessage = msg

       .showInput    = true
       .showError    = false

   end with ' }

end sub ' }
