Dim currentSlide as Slide
Dim shp as Shape

Set currentSlide = Application.ActiveWindow.View.Slide
For Each shp In currentSlide.Shapes
  ' Do something with the current shape referred to in variable 'shp'
  ' For example print the name of the shape in the Immediate Window
  Debug.Print shp.Name
Next shp
