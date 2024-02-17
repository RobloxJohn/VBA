Sub Change_All_To_LOWER_Case()

  Dim Rng As Range

  For Each Rng In Selection.Cells
   If Rng.HasFormula = False Then
     Rng.Value = LCase(Rng.Value)
   End If
  Next Rng

End Sub
