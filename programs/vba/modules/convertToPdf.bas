Attribute VB_Name = "convertToPdf"
' Tools > Reference Settings > Microsoft Word 16.0 Object Library
Option Explicit
Public Const debugFlag As Boolean = True

Public Sub ExecConvertToPdf()
    Dim file As Variant
    Dim paramList As collection
    Set paramList = CreateLatestFileList()
    Dim targetFolderName As String
    
    Dim param As Variant
    For Each param In paramList
        targetFolderName = param.Item(1)
        param.Remove 1
        Call ConvertToPdf(targetFolderName, param)
    Next param
    MsgBox "PDF conversion completed.", vbInformation

End Sub

Private Function CreateLatestFileList() As collection
    Dim myMap As collection
    Set myMap = New collection
    
    Dim isf_latestFile As collection
    Set isf_latestFile = New collection
    isf_latestFile.Add "ISMS（情報システム研究室）"
    isf_latestFile.Add "ISF01 "
    isf_latestFile.Add "ISF23-1 "
    isf_latestFile.Add "ISF25 "
    isf_latestFile.Add "ISF27-1 "
    isf_latestFile.Add "ISF27-2 "
    isf_latestFile.Add "ISF27-3 "
    isf_latestFile.Add "ISF27-4 "
    isf_latestFile.Add "ISF27-7 "
    myMap.Add isf_latestFile
    
    Dim qf_latestFile As collection
    Set qf_latestFile = New collection
    qf_latestFile.Add "QMS（情報システム研究室）"
    qf_latestFile.Add "QF01 "
    qf_latestFile.Add "QF04 "
    qf_latestFile.Add "QF06 "
    qf_latestFile.Add "QF13 "
    qf_latestFile.Add "QF23 "
    myMap.Add qf_latestFile
    
    Set CreateLatestFileList = myMap
End Function

Private Sub ConvertToPdf(targetFolderName As String, param As Variant)
    Dim wordObject As Object
    Dim fileName As String
    Dim fileExtension As String
    Dim wb As Workbook
    Dim wdDoc As Word.Document
    Dim fileCollection As collection
    Dim outputFilePath As String
    Dim nfcPath As String
    Dim targetFileCollection As collection
    Dim dummy As Variant
    Dim editPath As New ClassEditPath
    Dim inputFolderPath As String
    inputFolderPath = editPath.GetInputPath(targetFolderName)
    Dim outputFolderPath As String
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
                wb.Close SaveChanges:=False
            End If
        End If
        If fileExtension = "docx" Or fileExtension = "doc" Then
            Set wdDoc = wordObject.Documents.Open(path)
            If Not wdDoc Is Nothing Then
                wdDoc.ExportAsFixedFormat OutputFileName:=outputFilePath, ExportFormat:=17
                wdDoc.Close SaveChanges:=False
            End If
        End If
    Next path
    
    ' Clean up and close the application
    wordObject.Quit
    Set wordObject = Nothing
    
End Sub


