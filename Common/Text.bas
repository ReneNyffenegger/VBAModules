option explicit

public function regexpSplit(text as string, pattern as string) as string() ' {

  dim text_0 as string
  dim re     as new regExp

  re.pattern   = pattern
  re.global    = true
  re.multiLine = true

  text_0 = re.replace(text, vbNullChar)

  regexpSplit = strings.split(text_0, vbNullChar)

end function ' }

function lpad(text as String, length as integer, optional padChar as string = " ") ' {
    lpad = string(length - len(text), padChar) & text
end function ' }

function rpad(text as String, length as integer, optional padChar as string = " ") ' {
    rpad = text & string(length - len(text), padChar)
end function ' }

function parsePossibleDate(possibleDate as string) as variant ' {

    dim re as new regExp
    dim mc as     matchCollection

    re.pattern = "^(\d\d?)\.(\d\d?).(\d\d\d\d)( \d\d:\d\d)?$"

    set mc= re.execute(possibleDate)

    if mc.count > 0 then ' {
       parsePossibleDate = dateSerial(mc(0).subMatches(2), mc(0).subMatches(1), mc(0).subMatches(0))
       exit function
    end if ' }

    re.pattern = "^(\d\d\d\d)(\d\d)(\d\d)$"
    set mc= re.execute(possibleDate)

    if mc.count > 0 then ' {
       parsePossibleDate = dateSerial(mc(0).subMatches(0), mc(0).subMatches(1), mc(0).subMatches(2))
       exit function
    end if ' }

    parsePossibleDate = vbNull

end function ' }

sub test_parsePossibleDate() ' {

    if parsePossibleDate("28.08.2016 00:00") <> #2016-08-28# then ' {
       debug.print "Failed"
    end if ' }

end sub ' }
