option explicit

public function doesCollectionContain(col as collection, item as variant) ' {
'
' http://stackoverflow.com/a/991900/180275
'

  dim obj as variant

  on error got err

  obj = col(key)

  exit function


err:

  contains = false

end function ' }
