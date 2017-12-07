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
                   Cells(4, 10).Value = Port
                 Close #1
             ElseIf InStr(myFile.Name, "db.xml") Then
                
                DB = Left(myFile.Name, Len(myFile.Name) - 9)
                 Cells(5, 10).Value = DB
                
            End If
        Next
        Recurse = Recurse(mySubFolder.Path)
    Next

End Function

Sub RefreshSSASConnection()

Call Recurse("C:\Users\" & (Environ$("Username")) & "\AppData\Local\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\")

' developed by Matt Allington from http://exceleratorbi.com.au


    With ActiveWorkbook.Connections("PBID").OLEDBConnection
        .CommandText = Array("Model")
        .CommandType = xlCmdCube
        .Connection = Array( _
        "OLEDB;Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & DB & ";Data " _
        , _
        "Source=localhost:" & Port & ";MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2" _
        )
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .MaxDrillthroughRecords = 1000
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
        .RetrieveInOfficeUILang = True
    End With
    With ActiveWorkbook.Connections("PBID")
        .Name = "PBID"
        .Description = ""
    End With
    ActiveWorkbook.Connections("PBID").Refresh

End Sub
