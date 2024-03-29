Sub RemoveAnimationsFromAllSlides()
    Dim mySlide As slide
    Dim i As Long
 
    For Each mySlide In ActivePresentation.Slides
      For i = mySlide.TimeLine.MainSequence.Count To 1 Step -1
       'Remove Each Animation
       mySlide.TimeLine.MainSequence.Item(i).Delete
      Next i
    Next mySlide
     
End Sub
