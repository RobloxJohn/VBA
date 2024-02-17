Sub Delete_Whole_Folder()

 'You can use this to delete entire folder

 On Error Resume Next


 Kill "C:UsersAdmin_2.Dell-PcDesktopDelete Folder*.*"
 'Firstly it will delete all the files in the folder
 'Then below code will delete the entire folder if it is empty

 RmDir "C:UsersAdmin_2.Dell-PcDesktopDelete Folder"
 'Note: RmDir delete only a empty folder
 
 On Error GoTo 0

End Sub
