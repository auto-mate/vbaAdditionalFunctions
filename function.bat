Function VlookupEx(lookupVal,Lrange As Range,Offset)

  dim cellColl as Range

  VlookupEx=0
  for each cellColl in Lrange
    if cellColl.value=lookupVal then
        VlookupEx=cellColl.(0,Offset).value
      exit for
    endif
  next
  
End Function
