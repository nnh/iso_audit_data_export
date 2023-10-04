Attribute VB_Name = "convertToPdf"
Option Explicit

Public Sub ExecConvertToPdfLatestFile()
    Dim folderPathManager As New ClassFolderPathManager
    Dim latestFiles As Object
    Set latestFiles = folderPathManager.CreateLatestFileList()
    Dim latestFile As Variant
    Dim targetKeys As Variant
    targetKeys = latestFiles.Keys
    Dim targetKey As Variant
    For Each targetKey In targetKeys
        Call ConvertToPdf(folderPathManager.folderNameList(targetKey), latestFiles(targetKey))
    Next targetKey

End Sub

Private Sub ConvertToPdf(targetFolderName As String, param As Variant)
    Dim wordObject As Object
    Dim fileName As String
    Dim fileExtension As String
    Dim wb As Workbook
    Dim wdDoc As Word.Document
    Dim fileCollection As Collection
    Dim outputFilePath As String
    Dim nfcPath As String
    Dim targetFileCollection As Collection
    Dim dummy As Variant
    Dim editPath As New ClassEditPath
    Dim inputFolderPath As String
    Dim outputFolderPath As String
    inputFolderPath = editPath.GetInputPath(targetFolderName)
    outputFolderPath = editPath.GetOutputPath(targetFolderName)
    
    ' Create an instance of the application to manipulate files
    Set wordObject = CreateObject("Word.Application")
    ' Loop through all files in the input folder
    Set fileCollection = GetAllFilesInFolder(inputFolderPath)
    Set targetFileCollection = FilterFiles(fileCollection, param)
    
    Dim path As Variant
    For Each path In targetFileCollection
        fileName = GetFileName(path)
        fileExtension = GetFileExtension(fileName)
        outputFilePath = outputFolderPath & Replace(fileName, fileExtension, "pdf")
        ' Check if the file is either Excel or Word
        If fileExtension = "xlsx" Or fileExtension = "xls" Then
            ' Open the file
            Set wb = Workbooks.Open(fileName:=path)
            ' Check if the file is open
            If Not wb Is Nothing Then
                ' Save as PDF in the output folder
                wb.ExportAsFixedFormat Type:=xlTypePDF, _
                    fileName:=outputFilePath, _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
                ' Close the file without saving
                wb.Close savechanges:=False
            End If
        End If
        If fileExtension = "docx" Or fileExtension = "doc" Then
            Set wdDoc = wordObject.Documents.Open(path)
            If Not wdDoc Is Nothing Then
                wdDoc.ExportAsFixedFormat OutputFileName:=outputFilePath, ExportFormat:=17
                wdDoc.Close savechanges:=False
            End If
        End If
    Next path
    
    ' Clean up and close the application
    wordObject.Quit
    Set wordObject = Nothing
    
End Sub


