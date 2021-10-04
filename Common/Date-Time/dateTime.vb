option explicit


function dt_rfc_3339(dt as date) as string ' {
    dt_rfc_3339 = format(dt, "yyyy-mm-dd")
end function ' }


function dt_rfc_3339_sec(dt as date) as string ' {
    dt_rfc_3339_sec = format(dt, "yyyy-mm-dd hh:nn:ss")
end function ' }


function firstDayOfMonth(dt as date) as date ' {
    firstDayOfMonth = dateSerial(year(dt), month(dt), 1)
end function ' }


function lastDayOfMonth(dt as date) as date ' {
    lastDayOfMonth = dateAdd("m", 1, dateAdd("d", -1, firstDayOfMonth(dt)))
end function ' }
