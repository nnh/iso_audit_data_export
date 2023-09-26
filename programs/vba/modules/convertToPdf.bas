Attribute VB_Name = "convertToPdf"
' Tools > Reference Settings > Microsoft Word 16.0 Object Library
Option Explicit

Public Sub ConvertToPdf()
    Dim inputFolderPath As String
    Dim outputFolderPath As String
    Dim fileName As String
    Dim fileExtension As String
    Dim file As Variant
    Dim wordObject As Object
    Dim outputFilePath As String
    Dim wb As Workbook
    Dim wdDoc As Word.Document
    inputFolderPath = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\input\"
    outputFolderPath = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\output\"
    
    ' Create an instance of the application to manipulate files
    Set wordObject = CreateObject("Word.Application")
    
    ' Loop through all files in the input folder
    fileName = Dir(inputFolderPath & "*.*")
    Do While fileName <> ""
        ' Get the file extension
        fileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
        
        ' Check if the file is either Excel or Word
        If fileExtension = "xlsx" Or fileExtension = "xls" Then
            ' Open the file
            Set wb = Workbooks.Open(fileName:=inputFolderPath & fileName)
            
            ' Check if the file is open
            If Not wb Is Nothing Then
                ' Save as PDF in the output folder
                wb.ExportAsFixedFormat Type:=xlTypePDF, _
                    fileName:=outputFolderPath & Replace(fileName, fileExtension, "pdf"), _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
                ' Close the file without saving
                wb.Close SaveChanges:=False
            End If
        End If
        If fileExtension = "docx" Or fileExtension = "doc" Then
            Set wdDoc = wordObject.Documents.Open(inputFolderPath & fileName)
            outputFilePath = outputFolderPath & Replace(fileName, fileExtension, "pdf")
            wdDoc.ExportAsFixedFormat OutputFileName:=outputFilePath, ExportFormat:=17
            wdDoc.Close SaveChanges:=False
        End If
        ' Get the next file in the folder
        fileName = Dir
    Loop
    
    ' Clean up and close the application
    wordObject.Quit
    Set wordObject = Nothing
    
    MsgBox "PDF conversion completed.", vbInformation
End Sub



