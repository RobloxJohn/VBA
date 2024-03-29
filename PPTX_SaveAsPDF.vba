Sub SavePresentationAsPDF()
    Dim pptName As String
    Dim PDFName As String
    
    ' Save PowerPoint as PDF
    pptName = ActivePresentation.FullName
    ' Replace PowerPoint file extension in the name to PDF
    PDFName = Left(pptName, InStr(pptName, ".")) & "pdf"
    ActivePresentation.ExportAsFixedFormat PDFName, 2  ' ppFixedFormatTypePDF = 2
 
End Sub
