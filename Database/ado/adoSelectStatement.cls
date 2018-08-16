'  vim: ft=basic
'  
'  adoSelectStatement
'
option explicit

private cm            as ADODB.command

private sub class_initialize() ' {
    set cm          = new ADODB.command
    cm.commandType  = adCmdText
end sub ' }

public sub init(cn as ADODB.connection) ' {
     set cm.activeConnection = cn
end sub ' }

private sub class_terminate() ' {
    set cm = nothing
end sub ' }

public sub sqlFromFile(filename as string) ' {
    cm.commandText = slurpFile(filename)
end sub ' }

public sub addIntegerParameter(val as long)
    dim param as ADODB.parameter
    set param = cm.createParameter(, adInteger, adParamInput, , val)
    cm.parameters.append param
end sub

public function recordSet() as ADODB.recordSet ' {
    set recordSet = cm.execute
end function ' }
