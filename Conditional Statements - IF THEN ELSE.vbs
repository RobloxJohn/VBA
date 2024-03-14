Sub Macro1()
If Range("A1") > 100 Then
Range("B1").Value = 1
ElseIf Range("A1") > 50 Then
Range("B1") = 0.5
Else
Range("B1") = 0
End If
End Sub
