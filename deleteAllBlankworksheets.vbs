Sub DeleteBlankSheetsUsingCOUNTA()
    Dim ws As Worksheet
    Dim nonEmptyCells As Long
    
    ' Disable alerts to prevent confirmation boxes from appearing
    Application.DisplayAlerts = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Count non-empty cells in the worksheet using COUNTA function
        nonEmptyCells = Application.WorksheetFunction.CountA(ws.Cells)
        
        ' Delete the worksheet if COUNTA returns 0, meaning it's empty
        If nonEmptyCells = 0 Then
            ws.Delete
        End If
    Next ws
    
    ' Enable alerts
    Application.DisplayAlerts = True
End Sub
