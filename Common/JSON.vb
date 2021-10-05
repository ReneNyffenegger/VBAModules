'
'      V.1
'
option explicit

function json_val(val as variant) as string ' {

   select case varType(val) ' {
      case vbInteger, vbLong, vbSingle, vbDouble
           json_val = val

      case vbString
           json_val = """" & replace(val, """", "\""") & """"

      case vbDate
           json_val = """" & dt_iso_8601(val) & """" ' format(val, "yyyy-mm-dd""T""HH:MM:SS")

      case vbBoolean
           json_val = lcase(val)

      case vbEmpty
           json_val = "null"

      case else
           msgBox "json_val: Unrecognized datatype " & typeName(val) & ", " & varType(val)

      end select ' }

end function ' }

function json_key(name as variant) as string ' {
 '
 ' Create something that can be used as a JSON-Key, for example:
 '   "keyValue":

   json_key = json_val(name) & ":"
end function ' }

function json_name_value(name as string, val as variant) ' {
    json_name_value = json_val(name) & ": " & json_val(val)
end function ' }
