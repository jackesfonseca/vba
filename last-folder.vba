Function LastFolder() As String


    Dim i As Integer
    Dim oFSO As Object
    Dim oFolder As Object
    Dim objFile As Object
    Dim LastFolderInPath As String
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFolder = oFSO.GetFolder("C:\Users\B497834\OneDrive - IBERDROLA S.A\ARQUIVOS_REF")
    
    For Each objSubFolder In oFolder.subfolders
        If Left(objSubFolder.Name, 9) = "RELATORIO" And Len(objSubFolder.Name) = 25 Then
            LastFolderInPath = objSubFolder.Name
        End If
        
    Next objSubFolder
    
    LastFolder = LastFolderInPath
    
End Function
