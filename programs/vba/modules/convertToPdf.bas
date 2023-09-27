Attribute VB_Name = "convertToPdf"
' Tools > Reference Settings > Microsoft Word 16.0 Object Library
Option Explicit
Public Const outputFolderPath As String = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\output\"
Public Sub ExecConvertToPdf()
    Dim file As Variant
    Dim fiscalYear As Integer
    fiscalYear = GetFiscalYear() - 1
    Const inputParentPathHeader As String = "C:\Users\Mariko\Box\Projects\ISO\QMS・ISMS文書\04 記録\"
    Const inputParentPathFooter As String = "年度\ドラフト\"
    Dim inputParentPath As String
    inputParentPath = inputParentPathHeader & fiscalYear & inputParentPathFooter
    Dim inputFolderPath As String
    Dim targetFolderNames(1) As Variant
    targetFolderNames(0) = "ISMS（情報システム研究室）"
    targetFolderNames(1) = "QMS（情報システム研究室）"
    Dim targetFolderName As Variant
    For Each targetFolderName In targetFolderNames
        inputFolderPath = inputParentPath & targetFolderName & "\"
        Call ConvertToPdf(inputFolderPath)
    Next targetFolderName
    MsgBox "PDF conversion completed.", vbInformation

End Sub

Private Sub ConvertToPdf(inputFolderPath)
    Dim wordObject As Object
    Dim fileName As String
    Dim fileExtension As String
    Dim wb As Workbook
    Dim wdDoc As Word.Document
    Dim fileCollection As collection
    Dim outputFilePath As String
    Dim nfcPath As String
    
    ' Create an instance of the application to manipulate files
    Set wordObject = CreateObject("Word.Application")
    ' Loop through all files in the input folder
    Set fileCollection = GetAllFilesInFolder(inputFolderPath)
    Dim path As Variant
    For Each path In fileCollection
        fileName = Right(nfcPath, Len(nfcPath) - InStrRev(nfcPath, "\"))
        fileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
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
            wdDoc.ExportAsFixedFormat OutputFileName:=outputFilePath, ExportFormat:=17
            wdDoc.Close SaveChanges:=False
        End If
    Next path
    
    ' Clean up and close the application
    wordObject.Quit
    Set wordObject = Nothing
    
End Sub

