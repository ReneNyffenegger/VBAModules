option explicit

function collObjectOrNothing(coll as variant, byVal name as string) as object ' {
 '
 ' coll is defined as a variant so that all kinds of collections can
 ' be passed (such as excel.workbook.sheets etc.) as well as the vba.collection
 ' object
 '

   on error goto err_

      set collObjectOrNothing = coll.item(name)
      exit function

   err_:

'     if err.number <> 9 then ' 9 = Subscript out of range ' {
'        msgBox "collectionItemOrNothing: " & err.number & " - " & err.description
'     end if ' }

      set collObjectOrNothing = nothing

end function ' }

public function isKeyInColl(coll as variant, byVal key as variant) as boolean ' {
'
' For example in Excel:
'   debug.print(isKeyInColl(thisWorkbook.worksheets, "expected sheet name"))
'
'
' Compare
'    https://stackoverflow.com/a/991900/180275
'

  if collObjectOrNothing(coll, key) is nothing then
     isKeyInColl = false
  else
     isKeyInColl = true
  end if


'   dim obj as variant
'
'   on error goto err_
'
'   obj = col(key)
'
'   isKeyInColl = true
'
'   exit function
'
' err_:
'
'   isKeyInColl = false

end function ' }
