VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFunctionLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private paramList As Collection
Private folderNameList As Object
Private folderName As String
Public outputFolderPathList As Object
Const constIsms As String = "ISF"
Const constQms As String = "QF"

Private Sub Class_Initialize()
    Dim folderNameList As Object
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
    folderNameList.Add constIsms, "ISMS（情報システム研究室）"
    folderNameList.Add constQms, "QMS（情報システム研究室）"
    Set CreateTargetFolderNameList = folderNameList
End Function

Private Function CreateLatestFileDictionary(ByVal dic As Object, target() As String, key As String) As Object
    Dim latestFileList As Collection
    Set latestFileList = New Collection
    Dim fileName As Variant
    For Each fileName In target
        latestFileList.Add fileName
    Next fileName
    dic.Add key, latestFileList
    Set CreateLatestFileDictionary = dic

End Function



Private Function CreateLatestFileList() As Collection
    Dim myMap As Collection
    Set myMap = New Collection
    
    Dim isf_latestFile As Collection
    Set isf_latestFile = New Collection
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
    
    Dim qf_latestFile As Collection
    Set qf_latestFile = New Collection
    qf_latestFile.Add "QMS（情報システム研究室）"
    qf_latestFile.Add "QF01 "
    qf_latestFile.Add "QF04 "
    qf_latestFile.Add "QF06 "
    qf_latestFile.Add "QF13 "
    qf_latestFile.Add "QF23 "
    myMap.Add qf_latestFile
    
    Set CreateLatestFileList = myMap
End Function



Public Sub ConvertToPdfLatestFile()
    Dim file As Variant
    Dim targetFolderName As String
    
    Dim param As Variant
    For Each param In paramList
        targetFolderName = param.Item(1)
        param.Remove 1
        Call ConvertToPdf(targetFolderName, param)
    Next param
    MsgBox "PDF conversion completed.", vbInformation

End Sub
