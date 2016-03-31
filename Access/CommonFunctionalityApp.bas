sub appTitle(text as string) ' {

  call appProperty("AppTitle", dbText, text)
  application.refreshTitleBar

end sub ' }

function appProperty(strName As String, varType As Variant, varValue As Variant) As Integer ' {
' https://msdn.microsoft.com/en-us/library/bb256834(v=office.12).aspx
'
' TODO should this function be moved to its own «CommonFunctionalityApp.bas»?
'
  dim db As Object, prp As Variant
  const conPropNotFoundError = 3270

  set db = CurrentDb
  on error goto nok
  db.Properties(strName) = varValue
  appProperty = true

ok:
    exit function

nok:
  if err = conPropNotFoundError Then
     set prp = db.CreateProperty(strName, varType, varValue)
     db.Properties.Append prp
     resume
  else
     appProperty = false
     resume ok
  end if

end function ' }
