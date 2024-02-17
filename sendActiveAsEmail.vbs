Sub SendActiveSheetInEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim tempFilePath As String
    
    ' Create a new Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Create a new email item
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Save the active sheet to a temporary location
    tempFilePath = Environ("temp") & "\" & ActiveSheet.Name & ".xlsx"
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs Filename:=tempFilePath
    ActiveWorkbook.Close SaveChanges:=False
    
    ' Configure and send email
    With OutlookMail
        .To = "Excel@stevespreadsheetplanet.com"  ' Modify this line
        .Subject = "Latest Data Attached" ' Modify this line
        .Body = "Please find the attached the Latest Data." ' Modify this line
        .Attachments.Add tempFilePath ' Attach the temporary file
        .Send
    End With
    
    ' Delete the temporary file
    Kill tempFilePath
    
    ' Release the Outlook objects
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
