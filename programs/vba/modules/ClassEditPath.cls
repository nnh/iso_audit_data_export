VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEditPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private parentPathHeader As String
Private fiscalYear As Integer

Private Sub Class_Initialize()
    fiscalYear = GetFiscalYear()
    Dim parentPathHeaderHead As String
    parentPathHeaderHead = "C:\Users\Mariko\Box\Projects\ISO\QMS・ISMS文書\04 記録\"
    parentPathHeader = parentPathHeaderHead & fiscalYear & "年度\"
End Sub

Public Function GetInputPath(folderName As String) As String
    Dim targetPath As String
    targetPath = GetDraftPath(parentPathHeader)
    GetInputPath = targetPath & "\" & folderName & "\"
    '
End Function

Private Function GetDraftPath(parentFolderPath As String) As String
    Const constDraftTextPart As String = "ラフト"
    Dim parentFolder As Variant
    Dim fso As FileSystemObject
    Dim subFolders As Variant
    Dim subFolder As Variant
    Dim targetPath As String
    Set fso = New FileSystemObject
    Set parentFolder = fso.GetFolder(parentFolderPath)
    Set subFolders = parentFolder.subFolders
    targetPath = ""
    For Each subFolder In subFolders
        If InStr(1, subFolder.Name, constDraftTextPart) > 0 Then
            targetPath = subFolder.path
            Exit For
        End If
    Next subFolder
    GetDraftPath = targetPath

End Function

Public Function GetOutputPath(folderName As String) As String
    If debugFlag Then
        GetOutputPath = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\output\" & folderName & "\"
    Else
        GetOutputPath = parentPathHeader & "固定\" & folderName & "\"
    End If
End Function

Public Function GetDocumentListPath() As String
    Dim parentPathHeaderHead As String
    parentPathHeaderHead = "C:\Users\Mariko\Box\Projects\ISO\QMS・ISMS文書\"
    GetDocumentListPath = GetDraftPath(parentPathHeaderHead)

End Function
