'  vim: ft=basic
'
'  adoSelectStatement.cls
'
option explicit

private cm                 as ADODB.command
private rs_                as ADODB.recordSet
private firstCallOfRecord_ as boolean

private sub class_initialize() ' {
    dbg_.indent "adoSelectStatment: class_initialize"
    set cm             = new ADODB.command
    cm.commandType     = adCmdText
    firstCallOfRecord_ = true
    dbg_.dedent
end sub ' }

public sub init(cn as ADODB.connection) ' {
     dbg_.indent "adoSelectStatment: init"
     set cm.activeConnection = cn
     dbg_.dedent
end sub ' }

private sub class_terminate() ' {
    dbg_.indent "adoSelectStatment: class_terminate"
    set cm  = nothing
    set rs_ = nothing
    dbg_.dedent
end sub ' }

public sub sqlFromFile(filename as string) ' {
    cm.commandText = slurpFile(filename)
end sub ' }

public sub addIntegerParameter(val as long) ' {
    dbg_.indent "adoSelectStatment: addIntegerParameter, val = " & val
    dim param as ADODB.parameter
    set param = cm.createParameter(, adInteger, adParamInput, , val)
    cm.parameters.append param
    dbg_.dedent
end sub ' }

public sub addVarCharParameter(val as string) ' {
    dbg_.indent "adoSelectStatment: addVarCharParameter, val = " & val
    dim param as ADODB.parameter
    set param = cm.createParameter(, adVarChar, adParamInput, len(val), val)
    cm.parameters.append param
    dbg_.dedent
end sub ' }

public sub addDateParameter(val as date) ' {
    dbg_.indent "adoSelectStatment: addDateParementer, val = " & val
    dim param as ADODB.parameter
    set param = cm.createParameter(, adDate, adParamInput, , val)
    cm.parameters.append param
    dbg_.dedent
end sub ' }

public sub exec() ' as ADODB.recordSet  {
  on error goto err_
    dbg_.indent "adoSelectStatement.exec"
    set rs_   = cm.execute
    if rs_ is nothing then
       dbg_.text "rs_ is nothing"
    end if
    dbg_.dedent
    exit sub
  err_:
    msgBox "adoSelectStatement.exec: " & err.description
    dbg_.text "ERROR: " & err.description
    dbg_.dedent
end sub ' }

public function record() as boolean ' { End of records...
  on error goto err_
    dbg_.indent "adoSelectStatement.record"

    if rs_ is nothing then
       record = false
       dbg_.text "rs_ is nothing"
       dbg_.dedent
       exit function
    end if

    if not firstCallOfRecord_ then
       dbg_.text "Calling rs_.moveNext"
       rs_.moveNext
    else
       dbg_.text "Remember that we were already called"
       firstCallOfRecord_ = false
    end if

    if rs_.eof then
       dbg_.text "rs_.eof = true, will return false"
       record = false
    else
       dbg_.text "rs_.eof = false, will return true"
       record = true
    end if
    dbg_.dedent
    exit function
  err_:
    msgBox "adoSelectStatement, error: " & err.description
    dbg_.text "! ERROR: " & err.description
    dbg_.dedent
end function ' }

public function col(name as string) as ADODB.field ' {
    attribute name.vb_userMemId = 0
  on error goto err_:
    dbg_.indent "adoSelectStatement.col, name = " & name

    if rs_ is nothing then
       msgBox "adoSelectStatement: rs_ is nothing"
    end if
    dbg_.text "After rs_ check"

    if rs_.fields is nothing then
       msgBox "adoSelectStatement: rs_.fields is nothing"
    end if
    dbg_.text "After rs_.fields check"

    if rs_.fields(name) is nothing then
       msgBox "adoSelectStatement: rs_.fields(" & name & ") is nothing"
    end if
    dbg_.text "After rs_.fields(name) check"
    dbg_.text "rs_.fields(" & name & ") = " & rs_.fields(name)

    set col = rs_.fields(name)
    dbg_.dedent
    exit function
  err_:
    error_handler_("adoSelectStatement.col")
end function ' }
