Attribute VB_Name = "mGetAllFilesInFolder"
Option Explicit

Public Function GetAllFilesInFolder(ByVal folderPath As String) As collection
    Dim FSO As Object
    Dim parentFolder As Object
    Dim subfolder As Object
    Dim filesCollection As collection
    
    ' Create a FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the parent folder
    Set parentFolder = FSO.GetFolder(folderPath)
    
    ' Create a collection to store the file list
    Set filesCollection = New collection
    
    ' Add files in the parent folder to the collection
    AddFilesToCollection parentFolder, filesCollection
    
    ' Add files in subfolders to the collection as well (recursively)
    For Each subfolder In parentFolder.SubFolders
        AddFilesToCollection subfolder, filesCollection
    Next subfolder
    
    ' Return the collection of file paths
    Set GetAllFilesInFolder = filesCollection
End Function

Public Sub AddFilesToCollection(ByVal FOLDER As Object, ByRef filesCollection As collection)
    Dim file As Object
    
    ' Add files in the folder to the collection
    For Each file In FOLDER.Files
        filesCollection.Add file.path
    Next file
End Sub


