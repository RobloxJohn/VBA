Sub fill_array()

   Dim thisarray As Variant
   number_of_elements = 3     'number of elements in the array

   'must redim below to set size
   ReDim thisarray(1 To number_of_elements) As Integer
   'resizes this size of the array
   counter = 1
   fillmeup = 7
   For counter = 1 To number_of_elements
      thisarray(counter) = fillmeup
   Next counter

   counter = 1         'this loop shows what was filled in
   While counter <= UBound(thisarray)
      MsgBox thisarray(counter)
      counter = counter + 1
   Wend

End Sub
