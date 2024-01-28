Sub from_sheet_make_array()
   Dim thisarray As Variant
   thisarray = Range("a1:a10").Value

   counter = 1                'looping structure to look at array
   While counter <= UBound(thisarray)
      MsgBox thisarray(counter, 1)
      counter = counter + 1
   Wend
End Sub
