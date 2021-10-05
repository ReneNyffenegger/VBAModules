option explicit


function dt_rfc_3339(byVal dt as date) as string ' {
    dt_rfc_3339 = format(dt, "yyyy-mm-dd")
end function ' }

function dt_rfc_3339_sec(byVal dt as date) as string ' {
    dt_rfc_3339_sec = format(dt, "yyyy-mm-dd hh:nn:ss")
end function ' }

function dt_iso_8601(byVal dt as date) as string ' {
    dt_iso_8601 = format(dt, "yyyy-mm-dd\Thh:nn:ss")
end function ' }


function firstDayOfMonth(byVal dt as date) as date ' {
    firstDayOfMonth = dateSerial(year(dt), month(dt), 1)
end function ' }

function lastDayOfMonth(byVal dt as date) as date ' {
    lastDayOfMonth = dateAdd("m", 1, dateAdd("d", -1, firstDayOfMonth(dt)))
end function ' }
