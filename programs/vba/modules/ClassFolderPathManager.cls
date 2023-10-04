VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFolderPathManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private paramList As Collection
Public folderNameList As Object
Private folderName As String
Public outputFolderPathList As Object
Const constIsms As String = "ISF"
Const constQms As String = "QF"

Private Sub Class_Initialize()
    Set folderNameList = CreateTargetFolderNameList()
    Dim editPath As New ClassEditPath
    Dim folderName As Variant
    Dim filePath As String
    Set outputFolderPathList = CreateObject("Scripting.Dictionary")
    For Each folderName In folderNameList
        filePath = editPath.GetOutputPath(folderNameList(folderName))
        outputFolderPathList.Add folderName, filePath
    Next folderName
End Sub

Private Function CreateTargetFolderNameList() As Object
    Dim folderNameList As Object
    Set folderNameList = CreateObject("Scripting.Dictionary")
    folderNameList.Add constIsms, "ISMS（" & isr & "）"
    folderNameList.Add constQms, "QMS（" & isr & "）"
    Set CreateTargetFolderNameList = folderNameList
End Function

Public Function CreateLatestFileList() As Object
    Dim latestFiles As Object
    Set latestFiles = CreateObject("Scripting.Dictionary")
    
    Dim isf_latestFile As Collection
    Set isf_latestFile = New Collection
    isf_latestFile.Add "ISF01 "
    isf_latestFile.Add "ISF23-1 "
    isf_latestFile.Add "ISF25 "
    isf_latestFile.Add "ISF27-1 "
    isf_latestFile.Add "ISF27-2 "
    isf_latestFile.Add "ISF27-3 "
    isf_latestFile.Add "ISF27-4 "
    isf_latestFile.Add "ISF27-7 "
    latestFiles.Add constIsms, isf_latestFile
    
    Dim qf_latestFile As Collection
    Set qf_latestFile = New Collection
    qf_latestFile.Add "QF01 "
    qf_latestFile.Add "QF04 "
    qf_latestFile.Add "QF06 "
    qf_latestFile.Add "QF13 "
    qf_latestFile.Add "QF23 "
    latestFiles.Add constQms, qf_latestFile
    
    Set CreateLatestFileList = latestFiles
End Function

