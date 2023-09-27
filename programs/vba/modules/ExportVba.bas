Attribute VB_Name = "ExportVba"
' Tools > Reference Settings > Microsoft Scripting runtime
Option Explicit

Public Sub ExportVbaFiles()
    Dim exportPath As String
    Dim FSO As Scripting.FileSystemObject
    Dim exportModule As Variant
    exportPath = ThisWorkbook.path & "\modules"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(exportPath) Then
        FSO.CreateFolder exportPath
    End If
    For Each exportModule In ThisWorkbook.VBProject.VBComponents
        If exportModule.Type = 1 Then
            exportModule.Export exportPath & "\" & exportModule.Name & ".bas"
        End If
        If exportModule.Type = 2 Or exportModule.Type = 100 Then
            exportModule.Export exportPath & "\" & exportModule.Name & ".cls"
        End If
    Next exportModule
    
End Sub

