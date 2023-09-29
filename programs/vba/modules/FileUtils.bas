Attribute VB_Name = "FileUtils"
Option Explicit

Public Function FilterFiles(fileCollection As collection, param As Variant) As collection
    Dim filePath As Variant
    Dim targetFileName As Variant
    Dim fileName As String
    Dim latestFileName As String
    Dim tempFileName As Variant
    Dim j As Integer
    Dim removeTargetList() As String
    Dim removeCount As Integer
    Dim path As Variant
    Dim removePath As Variant
    Dim removeFlag As Boolean
    Dim targetFileCollection As collection
    removeCount = 0
            
    For Each targetFileName In param
        Dim target() As String
        Dim removeTargetPath() As String
        Dim i As Integer
        i = -1
        For Each filePath In fileCollection
            fileName = GetFileName(filePath)
            If Left(fileName, Len(targetFileName)) = targetFileName Then
                i = i + 1
                ReDim Preserve target(i)
                ReDim Preserve removeTargetPath(i)
                target(i) = fileName
                removeTargetPath(i) = filePath
            End If
        Next filePath
        If i > -1 Then
            latestFileName = GetLatestFileName(target)
            If latestFileName <> "" Then
                For j = LBound(removeTargetPath) To UBound(removeTargetPath)
                    If Right(removeTargetPath(j), Len(latestFileName)) <> latestFileName Then
                        ReDim Preserve removeTargetList(removeCount)
                        removeTargetList(removeCount) = removeTargetPath(j)
                        removeCount = removeCount + 1
                    End If
                Next j
            End If
        End If
    Next targetFileName
    
    Set targetFileCollection = New collection
    For Each path In fileCollection
        removeFlag = False
        For Each removePath In removeTargetList
            If path = removePath Then
                removeFlag = True
                Exit For
            End If
        Next removePath
        If Not (removeFlag) Then
            targetFileCollection.Add path
        End If
    Next path
    Set FilterFiles = targetFileCollection

End Function

Public Function GetLatestFileName(fileNameList() As String) As String
    Dim fileName As Variant
    Dim fileExtension As String
    Dim latestYmd() As Long
    ReDim latestYmd(UBound(fileNameList))
    Dim fileNameBodyList() As String
    ReDim fileNameBodyList(UBound(fileNameList))
    Dim i As Integer
    Dim maxValue As Long
    Dim targetFileName As String
    i = 0
    For Each fileName In fileNameList
        fileExtension = "." & GetFileExtension(CStr(fileName))
        fileNameBodyList(i) = Left(fileName, Len(fileName) - Len(fileExtension))
        latestYmd(i) = GetYmd(fileNameBodyList(i))
        i = i + 1
    Next fileName
    maxValue = GetMaxValue(latestYmd)
    i = GetTargetFileNameIndex(maxValue, fileNameBodyList)
    If i > -1 Then
        targetFileName = fileNameList(i)
    Else
        targetFileName = ""
    End If

    GetLatestFileName = targetFileName
End Function

Private Function GetTargetFileNameIndex(ymd As Long, fileNameBodyList() As String) As Integer
    Dim strMaxValue As String
    Dim fileNameBody As Variant
    Dim targetFileName As String
    Dim i As Integer
    If ymd < 1000000 Then
        strMaxValue = "(" & CStr(ymd) & ")"
    Else
        strMaxValue = CStr(ymd)
    End If
    i = 0
    For Each fileNameBody In fileNameBodyList
        If Right(fileNameBody, Len(strMaxValue)) = strMaxValue Then
             GetTargetFileNameIndex = i
             Exit Function
        End If
        i = i + 1
    Next fileNameBody
    GetTargetFileNameIndex = -1

End Function

Public Function GetMaxValue(arr() As Long) As Long
    Dim maxVal As Long
    Dim i As Long
    
    If UBound(arr) < LBound(arr) Then
        GetMaxValue = 0
        Exit Function
    End If
    
    maxVal = arr(LBound(arr))
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) > maxVal Then
            maxVal = arr(i)
        End If
    Next i
    
    GetMaxValue = maxVal
End Function


Public Function GetYmd(fileName As String) As Long
    Dim regex As RegExp
    Set regex = New RegExp
    regex.Pattern = "\(\d{6}\)$|\d{8}$"
    Dim tempFileName As String
    If regex.test(fileName) Then
        tempFileName = CLng(regex.Execute(fileName)(0).Value)
        If tempFileName < 0 Then
            tempFileName = tempFileName * -1
        End If
    Else
        tempFileName = -1
    End If
    GetYmd = tempFileName
End Function

Public Function GetAllFilesInFolder(ByVal folderPath As String) As collection
    Dim fso As Object
    Dim parentFolder As Object
    Dim subFolder As Object
    Dim filesCollection As collection
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the parent folder
    Set parentFolder = fso.GetFolder(folderPath)
    
    ' Create a collection to store the file list
    Set filesCollection = New collection
    
    ' Add files in the parent folder to the collection
    AddFilesToCollection parentFolder, filesCollection
    
    ' Add files in subfolders to the collection as well (recursively)
    For Each subFolder In parentFolder.subFolders
        AddFilesToCollection subFolder, filesCollection
    Next subFolder
    
    ' Return the collection of file paths
    Set GetAllFilesInFolder = filesCollection
End Function

Public Sub AddFilesToCollection(ByVal FOLDER As Object, ByRef filesCollection As collection)
    Dim file As Object
    
    ' Add files in the folder to the collection
    For Each file In FOLDER.Files
        filesCollection.Add file.path
    Next file
End Sub

Public Function GetFileName(path As Variant) As String
    GetFileName = Right(path, Len(path) - InStrRev(path, "\"))
End Function

Public Function GetFileExtension(fileName As String)
    GetFileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
End Function
