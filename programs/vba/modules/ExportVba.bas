Attribute VB_Name = "ExportVba"
Option Explicit

Public Sub ExportVbaFiles()
    Dim exportPath As String
    Dim fso As Scripting.FileSystemObject
    Dim exportModule As Variant
    exportPath = ThisWorkbook.path & "