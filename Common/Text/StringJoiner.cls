' vi: ft=vb
'
' V0.1

option explicit

private first   as boolean
private joiner  as string
private buf     as stringBuffer

public sub init(joiner_ as string, optional bufSize as long = 1000) ' {
    first  = true
    joiner = joiner_

    set buf = new stringBuffer : buf.init(bufSize)
end sub ' }

function add(txt as variant) as string ' {

    if first then
       buf.append(txt)
       first = false
    else
       buf.append(joiner & txt)
    end if

end function ' }

function value() as string ' {
     attribute value.vb_userMemId = 0
     value = buf
end function ' }
