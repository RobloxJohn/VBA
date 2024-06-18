Sub shufflerange()
Dim Iupper As Integer
Dim Ilower As Integer
Dim Ifrom As Integer
Dim Ito As Integer
Dim i As Integer
Iupper = InputBox("What is the highest slide number to shuffle")
Ilower = InputBox("What is the lowest slide number to shuffle")
If Iupper > ActivePresentation.Slides.Count Or Ilower < 1 Then GoTo err
For i = 1 To 2*Iupper
Randomize
Ifrom = Int((Iupper - Ilower + 1) * Rnd + Ilower)
Ito = Int((Iupper - Ilower + 1) * Rnd + Ilower)
ActivePresentation.Slides(Ifrom).MoveTo (Ito)
Next i
Exit Sub
err:
MsgBox "Your choices are out of range", vbCritical
End Sub
