Sub BackupWorkbookWithTimestamp()
    Dim filePath As String
    Dim fileName As String
    Dim timeStamp As String
    
    ' Create a timestamp in the format "yyyyMMdd_hhmmss"
    timeStamp = Format(Now, "yyyyMMdd_hhmmss")
    
    ' Get the current workbook's path and name
    filePath = ThisWorkbook.Path
    fileName = ThisWorkbook.Name
    
    ' Create the new name for the backup file by appending the timestamp
    Dim backupFileName As String
    backupFileName = Replace(fileName, ".xlsx", "_" & timeStamp & ".xlsx")
    
    ' Generate the full path for the backup file
    Dim backupFilePath As String
    backupFilePath = filePath & "\" & backupFileName
    
    ' Save a backup copy of the current workbook
    ThisWorkbook.SaveCopyAs backupFilePath
End Sub
