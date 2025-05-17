Dim currentSlide as Slide
Dim shp as Shape

Set currentSlide = Application.ActiveWindow.View.Slide
For Each shp In currentSlide.Shapes
  ' Check if the shape type is msoTextBox 
  If shp.Type = 17 Then ' msoTextBox = 17
    'Print the text in the TextBox
    Debug.Print shp.TextFrame2.TextRange.Text
  End If
Next shp
