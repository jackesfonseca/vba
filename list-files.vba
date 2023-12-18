Sub ListFiles()

    Dim i As Integer
    Dim oFSO As Object
    Dim oFolder As Object
    Dim objFile As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFolder = oFSO.GetFolder("C:\Users\test")
    
    For Each objFile In oFolder.Files
        Cells(i + 1, 1) = objFile.Name
        i = i + 1
    Next objFile

End Sub
