'  vim: ft=basic
'
'  adoStatement.cls
'
option explicit

private cm                 as ADODB.command
private rs_                as ADODB.recordSet
private firstCallOfRecord_ as boolean

private sub class_initialize() ' {
    set cm             = new ADODB.command
    cm.commandType     = adCmdText
end sub ' }

public sub init(cn as ADODB.connection) ' {
     set cm.activeConnection = cn
end sub ' }

private sub class_terminate() ' {
    set cm  = nothing
    set rs_ = nothing
end sub ' }

public sub sql(text as string) ' {
    cm.commandText = text
end sub ' }

public sub sqlFromFile(filename as string) ' {
    cm.commandText = slurpFile(filename)
end sub ' }

public sub defineParameters(paramArray types()) ' {

    dim param   as ADODB.Parameter
    dim paramNo as long
    
    for paramNo = lBound(types) to uBound(types) ' {
    
        dim type_ as long
        type_ = types(paramNo)
        
        if type_ <> adVarChar then
           set param = cm.createParameter(, type_, adParamInput)
        else
         '
         ' Increase paramNo by one to access the required length for
         ' the adVarChar type when creating the parameter:
         '
           paramNo = paramNo + 1
           set param = cm.createParameter(, type_, adParamInput, types(paramNo))
        end if

        cm.parameters.append param 
    next paramNo ' }

end sub ' }

public sub addIntegerParameter(val as long) ' {
    dim param as ADODB.parameter
    set param = cm.createParameter(, adInteger, adParamInput, , val)
    cm.parameters.append param
end sub ' }

public sub addVarCharParameter(val as string) ' {
    dim param as ADODB.parameter
    set param = cm.createParameter(, adVarChar, adParamInput, len(val), val)
    cm.parameters.append param
end sub ' }

public sub addDateParameter(val as date) ' {
    dim param as ADODB.parameter
    set param = cm.createParameter(, adDate, adParamInput, , val)
    cm.parameters.append param
end sub ' }

public sub exec(paramArray paramValues()) ' {
 on error goto err_
    dim paramNo as long
  '
  ' I keep forgetting if arrays are zero or one based in VBA
  ' Therefore, make it explicit by declareing an additional, possibly
  ' redundant variable (i):
  '  
    dim i       as long
    
    for paramNo = lBound(paramValues) to uBound(paramValues) ' {
    
        cm.parameters(i) = paramValues(paramNo)
        i = i + 1

    next paramNo ' }
    
    set rs_            = cm.execute
    firstCallOfRecord_ =  true
    exit sub

  err_:
    dbg_.text("adoSelectStatment.exec: " & err.description)
    err.raise err.number, err.source, err.description
end sub ' }

public function record() as boolean ' {
  on error goto err_

    if rs_ is nothing then
       record = false
       exit function
    end if

    if not firstCallOfRecord_ then
       rs_.moveNext
    else
     ' Remember that we were already called
       firstCallOfRecord_ = false
    end if

    if rs_.eof then
       record  = false
       set rs_ = nothing
    else
       record = true
    end if
    exit function
  err_:
    dbg_.text "adoStatement.record: " & err.description
    err.raise err.number, err.source, err.description
end function ' }

public function col(name as string) as ADODB.field ' {
    attribute name.vb_userMemId = 0
  '
  ' The name.vb_userMemId = 0 specifies this function to be
  ' the default function.
  '
    if rs_ is nothing then
       msgBox "adoStatement: rs_ is nothing"
    end if

    if rs_.fields is nothing then
       msgBox "adoStatement: rs_.fields is nothing"
    end if

    if rs_.fields(name) is nothing then
       msgBox "adoSStatement: rs_.fields(" & name & ") is nothing"
    end if

    set col = rs_.fields(name)
end function ' }
