option explicit
'
'  V0.3
'
'  Add reference to Regular Expression Library:
'      call application.VBE.activeVBProject.references.addFromGuid("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5,  5)
'

public function regexpSplit(text as string, pattern as string) as string() ' {

  dim text_0 as string
  dim re     as new regExp

  re.pattern   = pattern
  re.global    = true
  re.multiLine = true

  text_0 = re.replace(text, vbNullChar)

  regexpSplit = strings.split(text_0, vbNullChar)

end function ' }

function lpad(byVal text as string, length as integer, optional byVal padChar as string = " ") ' {
    lpad = string(length - len(text), padChar) & text
end function ' }

function rpad(byVal text as string, length as integer, optional byVal padChar as string = " ") ' {
    rpad = text & string(length - len(text), padChar)
end function ' }

function parsePossibleDate(possibleDate as variant) as variant ' {

    dim re as new regExp
    dim mc as     matchCollection

    re.ignorecase = true

    if isEmpty(possibleDate) then ' {
       parsePossibleDate = cvDate(null)
       exit function
    end if ' }

    if isError(possibleDate) then ' {
       parsePossibleDate = cvDate(null)
       exit function
    end if ' }

    if possibleDate = "0" then ' {
    '
    '  Probably not intended to have december 30th 1899 as date.
    '
       parsePossibleDate = cvDate(null)
       exit function
    end if ' }

    re.pattern = "^(\d\d?)\.(\d\d?).(\d\d\d\d)( \d\d:\d\d)?$" ' {

    set mc= re.execute(possibleDate)

    if mc.count > 0 then ' {
       parsePossibleDate = dateSerial(mc(0).subMatches(2), mc(0).subMatches(1), mc(0).subMatches(0))
       exit function
    end if ' }

    ' }

    re.pattern = "^(\d\d\d\d)(\d\d)(\d\d)$" ' {
    set mc= re.execute(possibleDate)

    if mc.count > 0 then ' {
       parsePossibleDate = dateSerial(mc(0).subMatches(0), mc(0).subMatches(1), mc(0).subMatches(2))
       exit function
    end if ' }

    ' }

    re.pattern = "^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec) (\d+), (\d{4})$" ' {
    set mc= re.execute(possibleDate)

    if mc.count > 0 then ' {

       dim m as long

       select case lCase(mc(0).subMatches(0)) ' {
          case "jan" : m = 1
          case "feb" : m = 2
          case "mar" : m = 3
          case "apr" : m = 4
          case "may" : m = 5
          case "jun" : m = 6
          case "jul" : m = 7
          case "aug" : m = 8
          case "sep" : m = 9
          case "oct" : m =10
          case "nov" : m =11
          case "dec" : m =12
       end select ' }

       parsePossibleDate = dateSerial(mc(0).subMatches(2), m , mc(0).subMatches(1))
       exit function
    end if ' }

    ' }

    re.pattern = "^(\d+)$" ' The »date« might just be the numbers since 1899-12-30. ' {
    set mc = re.execute(possibleDate)

    if mc.count > 0 then ' {
       parsePossibleDate = cDate(mc(0).subMatches(0))
       exit function
    end if ' }

    ' }

    parsePossibleDate = cvDate(null)

end function ' }

sub test_parsePossibleDate() ' {

    if parsePossibleDate("28.08.2016 00:00") <> #2016-08-28# then ' {
       debug.print "Failed"
    end if ' }

    if parsePossibleDate("dec 17, 2019") <> #2019-12-17# then ' {
       debug.print "Failed"
    end if ' }

end sub ' }
