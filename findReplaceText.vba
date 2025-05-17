Sub FindAndReplaceText()
    Dim mySlide As slide
    Dim shp As Shape
    Dim findWhat As String
    Dim replaceWith As String
    Dim ShpTxt As TextRange
    Dim TmpTxt As TextRange

    findWhat = "jackal"
    replaceWith = "fox"
     
    ' Find and Find and Replace
    For Each mySlide In ActivePresentation.Slides
      For Each shp In mySlide.Shapes
        If shp.Type = 17 Then ' msoTextBox = 17
          Set ShpTxt = shp.TextFrame.TextRange
          'Find First Instance of "Find" word (if exists)
          Set TmpTxt = ShpTxt.Replace(findWhat, _
             Replacewhat:=replaceWith, _
             WholeWords:=True)
     
          'Find Any Additional instances of "Find" word (if exists)
          Do While Not TmpTxt Is Nothing
            Set ShpTxt = ShpTxt.Characters(TmpTxt.Start + TmpTxt.Length, ShpTxt.Length)
            Set TmpTxt = ShpTxt.Replace(findWhat, _
              Replacewhat:=replaceWith, _
              WholeWords:=True)
          Loop
        End If
      Next shp
    Next mySlide
 
End Sub
