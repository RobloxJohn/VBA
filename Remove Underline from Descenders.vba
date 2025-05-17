Sub RemoveUnderlineFromDescenders()
    Dim mySlide As slide
    Dim shp As Shape
    Dim descenders_list As String
    Dim phrase As String
    Dim x As Long
    
    ' Remove underlines from Descenders
    descenders_list = "gjpqy"
    For Each mySlide In ActivePresentation.Slides
      For Each shp In mySlide.Shapes
        If shp.Type = 17 Then ' msoTextBox = 17
         ' Remove underline from letters "gjpqy"
         With shp.TextFrame.TextRange
            phrase = .Text
           For x = 1 To Len(.Text)
             If InStr(descenders_list, Mid$(phrase, x, 1)) > 0 Then
              .Characters(x, 1).Font.Underline = False
             End If
           Next x
         End With
       End If
      Next shp
    Next mySlide
 
End Sub
