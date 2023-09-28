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

Private Sub Class_Initialize()
    Dim fiscalYear As Integer
    fiscalYear = GetFiscalYear()
    Dim parentPathHeaderHead As String
    parentPathHeaderHead = "C:\Users\Mariko\Box\Projects\ISO\QMS・ISMS文書\04 記録\"
    parentPathHeader = parentPathHeaderHead & fiscalYear & "年度\"
End Sub

Public Function GetInputPath(folderName As String) As String
    Const constDraftText1 As String = "ドラフト\"
    Const constDraftText2 As String = "ト゛ラフト\"
    Const constDraftTextPart As String = "ラフト"
    Dim inputPath As String
    Dim fso As FileSystemObject
    Dim parentFolder As Variant
    Dim subFolders As Variant
    Dim targetPath As String
    targetPath = ""
    Set fso = New FileSystemObject
    Set parentFolder = fso.GetFolder(parentPathHeader)
    Set subFolders = parentFolder.subFolders
    Dim subFolder As Variant
    For Each subFolder In subFolders
        If Right(subFolder.Name, Len(constDraftTextPart)) = constDraftTextPart Then
            targetPath = subFolder.path & "\" & folderName & "\"
            Exit For
        End If
    Next subFolder
    GetInputPath = targetPath
    '
End Function

Public Function GetOutputPath(folderName As String) As String
    If debugFlag Then
        GetOutputPath = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\output\" & folderName & "\"
    Else
        GetOutputPath = parentPathHeader & "固定\" & folderName & "\"
    End If
End Function

