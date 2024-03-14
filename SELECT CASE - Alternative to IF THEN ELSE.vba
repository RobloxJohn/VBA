Sub Macro11()
Select Case Range("A1").Value
   Case Is > 100
    Range("B1").Value = 1
   Case Is > 50
   Range("B1").Value = 0.5
   Case Else
   Range("B1").Value = 0
   End Select
End Sub 
