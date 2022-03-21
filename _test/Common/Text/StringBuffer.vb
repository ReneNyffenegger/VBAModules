option explicit

sub testStringBuffer () ' {

    dim sb as new StringBuffer : sb.init 5

    sb.append "foo"
    sb.append ", bar"

    if sb <> "foo, bar" then
       msgBox "Expected foo, bar, but got: " & sb
    end if

    sb.append " and baz. And quite a lot more!"

    if sb.value <> "foo, bar and baz. And quite a lot more!" then ' {
       msgBox sb.value
    end if ' }

  '
  ' StringBuffer V0.2: test default attribute:
  '
    if sb <> "foo, bar and baz. And quite a lot more!" then ' {
       msgBox "Accessing default attribute did not succeed"
    end if ' }

    sb.append("<")
    if sb <> "foo, bar and baz. And quite a lot more!<" then ' {
       msgBox "trailing < not found"
    end if ' }

    timeIt

end sub ' }

private sub timeIt ' {

    dim t0 as double

    t0 = timer

    dim i   as long
    dim str as string
    for i = 1 to 25000
        str = str & "abcdefghijklmnopqrstuvwxzy"
    next i

    debug.print "time string      : " & format((timer - t0) / 86400, "hh:mm:ss")

    t0 = timer
    dim strBuf as new stringBuffer : strBuf.init (10000& * 26)
    for i = 1 to 25000
        strBuf.append "abcdefghijklmnopqrstuvwxzy"
    next i

    debug.print "time stringBuffer: " & format((timer - t0) / 86400, "hh:mm:ss")

    if str <> strBuf.value then
       msgBox "str <> strBuf.value"
    end if

end sub ' }
