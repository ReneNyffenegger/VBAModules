' vi: ft=vb
'
' V0.2 

option explicit

private buf     as string
private cur_len as long

public sub init(size as long) ' {

    buf     = space$(size)
    cur_len = 0

end sub ' }

public sub append(text as string) ' {

    dim next_len as long
    next_len = cur_len + len(text)

    if len(buf) < next_len then
       dim cur_text as string
       cur_text = value

       buf = space$(len(buf) * 2)
       cur_len = 0
       append(cur_text)
    end if

    mid$(buf, 1+cur_len, next_len) = text
    cur_len = next_len

end sub ' }

public function value as string ' {
    attribute value.vb_UserMemId = 0

    value = mid$(buf, 1, cur_len)
end function ' }
