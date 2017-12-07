Global Port As String
Global DB As String

Function Recurse(sPath As String) As String

    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim mySubFolder As Folder
    Dim myFile As File

    Set myFolder = FSO.GetFolder(sPath)

    For Each mySubFolder In myFolder.SubFolders
         For Each myFile In mySubFolder.Files
            If InStr(myFile.Name, "msmdsrv.port.txt") Then
                  Open myFile.Path For Input As #1
                   Port = Input$(5, 1)
                 
                 Close #1
             ElseIf InStr(myFile.Name, "db.xml") Then
             Debug.Print InStr(myFile.Name, ".")
                
                DB = Left(myFile.Name, InStr(myFile.Name, ".") - 1)
                 
                
            End If
        Next
        Recurse = Recurse(mySubFolder.Path)
    Next

End Function



Call Recurse("C:\Users\" & (Environ$("Username")) & "\AppData\Local\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\")
