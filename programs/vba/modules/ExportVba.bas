Attribute VB_Name = "ExportVba"
Option Explicit

Public Sub ExportVbaFiles()
    Dim exportPath As String
    Dim fso As Scripting.FileSystemObject
    Dim exportModule As Variant
    exportPath = ThisWorkbook.Path & "\modules"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
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

