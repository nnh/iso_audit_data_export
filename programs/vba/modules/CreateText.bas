Attribute VB_Name = "CreateText"
Option Explicit
Public Sub ExecCreateTextFunction()
    Call GetTargetWorksheets
End Sub
Private Sub GetTargetWorksheets()
    Dim wb As Workbook
    Dim i As Integer
    Dim targetSheetName(1) As String
    targetSheetName(0) = "文書管理台帳(2)"
    targetSheetName(1) = "文書管理台帳(3)"
    Dim targetSheetValues As Variant
    Set wb = GetTargetDocument()
    If wb Is Nothing Then
        Exit Sub
    End If
    Dim tempTargetSheetValues As Variant
    Dim addRowCount As Integer
    Dim startRow As Integer
    Dim inputLastRow As Integer
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim wk1Worksheet As Worksheet
    Set wk1Worksheet = ThisWorkbook.Worksheets("wk1")
    wk1Worksheet.Cells.Clear
    startRow = 1
    Dim tempAddress As String
    For i = LBound(targetSheetName) To UBound(targetSheetName)
        tempAddress = wb.Worksheets(targetSheetName(i)).UsedRange.Address(ReferenceStyle:=xlR1C1)
        inputLastRow = Split(Split(tempAddress, "R")(2), "C")(0)
        tempTargetSheetValues = wb.Worksheets(targetSheetName(i)).Range("A1:J" & inputLastRow).Value
        wk1Worksheet.Range("A" & startRow & ":J" & startRow + inputLastRow - 1).Value = tempTargetSheetValues
        startRow = inputLastRow + 1
    Next i
    wb.Close savechanges:=False
    Dim targetValuesList() As String
    Dim targetValues() As String
    Dim fileNames() As String
    targetSheetValues = wk1Worksheet.UsedRange.Value
    targetValues = GetTargetValues(targetSheetValues)
    fileNames = CreateFileNames(targetValues)
    Dim fileNameAndTextDic As Object
    Set fileNameAndTextDic = CreateObject("Scripting.Dictionary")
    Dim fileName As Variant
    For Each fileName In fileNames
        fileNameAndTextDic.Add fileName, ""
    Next fileName
    fileNameAndTextDic.Add "ISF19 仕様書", "Box/Datacenter/ISR/Ptosh/Ptosh Validation"
    fileNameAndTextDic.Add "ISF22 ログデータ DC入退室", "\\aronas\Archives\Log\DC入退室"
    fileNameAndTextDic.Add "ISF22 ログデータ PivotalTracker", "\\aronas\Archives\PivotalTracker"
    fileNameAndTextDic.Add "ISF22 ログデータ UTM", "\\aronas\Archives\ISR\SystemAssistant\monthlyOperations"
    fileNameAndTextDic.Add "ISF22 ログデータ VPN", "\\aronas\Archives\Log\VPN"
    Dim qf20Path As String
    qf20Path = "\Box\Projects\ISO\QMS・ISMS文書\06 その他\研修資料\" & GetFiscalYear() & "年度 "
    fileNameAndTextDic.Add "QF30 教育資料", qf20Path
    Call ExecCreateTextFile(fileNameAndTextDic)
    wk1Worksheet.Cells.Clear
        
End Sub
Private Sub ExecCreateTextFile(fileNameAndTextDic As Object)
    Dim folderPathManager As New ClassFolderPathManager
    Dim pathList As Object
    Set pathList = folderPathManager.outputFolderPathList
    Dim pathKeys As Variant
    pathKeys = pathList.Keys
    Dim fileName As Variant
    Dim fileHead As Variant
    Dim filePath As String
    Dim fileNames As Variant
    fileNames = fileNameAndTextDic.Keys
    For Each fileName In fileNames
        Dim text As String
        text = fileNameAndTextDic(fileName)
        For Each fileHead In pathKeys
            If Left(fileName, Len(fileHead)) = fileHead Then
                filePath = pathList(fileHead)
                Call CreateTextFile(filePath, CStr(fileName), text)
                Exit For
            End If
        Next fileHead
    Next fileName
End Sub
Private Function CreateFileNames(targetValues() As String) As String()
    Dim index As Object
    Dim fileName As String
    Set index = CreateTargetValuesIndex()
    Dim i As Integer
    Dim arr() As String
    ReDim arr(UBound(targetValues, 1))
    Dim targetFilenames() As String
    Dim fileNameCount As Integer
    fileNameCount = 0
    For i = LBound(targetValues, 2) To UBound(targetValues, 2)
        Dim j As Integer
        For j = LBound(targetValues, 1) To UBound(targetValues, 1)
            arr(j) = targetValues(j, i)
        Next j
        fileName = CreateFileName(arr, index, targetValues)
        If fileName <> "" Then
            ReDim Preserve targetFilenames(fileNameCount)
            targetFilenames(fileNameCount) = fileName
            fileNameCount = fileNameCount + 1
        End If
    Next i
    CreateFileNames = targetFilenames
End Function
Private Function CreateFileName(values() As String, index As Object, targetValues() As String) As String
    Const constRef As String = "参照"
    Dim textHeader As String
    Dim temp As Variant
    Dim targetDept As String
    Const dc As String = "データ管理室"
    If (values(index("isr")) = "○" And _
        ( _
         Left(values(index("category")), 5) <> "ISF12" _
        ) _
       ) Or _
       Left(values(index("category")), 4) = "QF30" Then
        CreateFileName = ""
        Exit Function
    End If
    ' 「【区分】【記録名】」までは共通
    textHeader = values(index("category")) & " " & values(index("itemName"))
    Dim fileName As String
    Dim refCategoryText As String
    fileName = textHeader
    If InStr(1, values(index("format")), "紙") And _
       ( _
        Left(values(index("category")), 4) <> "QF22" _
       ) Then
        targetDept = dc
        refCategoryText = "紙保管"
    ' ISRに…参照の文字列が存在する場合、ファイル名は「【区分】【記録名】情報システム研究室【区分】参照」
    ElseIf InStr(1, values(index("isr")), constRef) > 0 Then
        targetDept = isr
        refCategoryText = GetRefCategoryByItemName(values(index("category")), values(index("itemName")), index, targetValues)
        refCategoryText = refCategoryText & constRef
    ' DCに…参照の文字列が存在する場合、ファイル名は「【区分】【記録名】データ管理室【区分】参照」
    ElseIf InStr(1, values(index("dc")), constRef) > 0 Then
        targetDept = dc
        refCategoryText = GetRefCategoryByItemName(values(index("category")), values(index("itemName")), index, targetValues)
        refCategoryText = refCategoryText & constRef
    ' ISRが空白でDCが○の場合、ファイル名は「【区分】【記録名】データ管理室【区分】参照」
    ElseIf values(index("isr")) = "" And values(index("dc")) <> "" Then
        targetDept = dc
        refCategoryText = values(index("category"))
        refCategoryText = refCategoryText & constRef
    End If
    fileName = fileName & " " & targetDept & refCategoryText
    fileName = Replace(fileName, vbLf, "")
    
    CreateFileName = fileName

End Function
Private Function GetRefCategoryByItemName(category As String, itemName As String, index As Object, targetValues() As String) As String
    Dim refCategory As String
    Dim i As Integer
    For i = LBound(targetValues, 2) To UBound(targetValues, 2)
        If Replace(targetValues(index("itemName"), i), vbLf, "") = itemName And targetValues(index("category"), i) <> category Then
            refCategory = targetValues(index("category"), i)
            Exit For
        End If
    Next i
    GetRefCategoryByItemName = refCategory

End Function
Private Function GetRefCategoryName(inputText As String, constRef As String) As String
    Dim splitByLf As Variant
    splitByLf = Split(inputText, vbLf)
    Dim tempValue As Variant
    For Each tempValue In splitByLf
        If InStr(1, tempValue, constRef) > 0 Then
            GetRefCategoryName = Replace(tempValue, constRef, "")
            Exit Function
        End If
    Next tempValue
    GetRefCategoryName = ""
End Function
Private Function CreateTargetValuesIndex() As Object
    Dim docIndex As Object
    Set docIndex = getDocumentColIndex()
    Dim docIndexCount As Integer
    docIndexCount = docIndex.count
    Dim docIndexKeys As Variant
    docIndexKeys = docIndex.Keys
    Dim i As Integer
    Dim tempKey As Variant
    i = 0
    Dim newIndex As Object
    Set newIndex = CreateObject("Scripting.Dictionary")
    For Each tempKey In docIndexKeys
        newIndex.Add tempKey, i
        i = i + 1
    Next tempKey
    Set CreateTargetValuesIndex = newIndex
End Function
Private Function getDocumentColIndex() As Object
    Dim index As Object
    Set index = CreateObject("Scripting.Dictionary")
    index.Add "category", 1
    index.Add "itemName", 2
    index.Add "format", 4
    index.Add "dc", 9
    index.Add "isr", 10
    Set getDocumentColIndex = index
End Function
Private Function GetCategoryHeaderList() As String()
    Dim targetCategoryHeader(1) As String
    targetCategoryHeader(0) = "QF"
    targetCategoryHeader(1) = "ISF"
    GetCategoryHeaderList = targetCategoryHeader
End Function
Private Function GetTargetValues(inputValues As Variant) As String()
    Dim targetCategoryHeader() As String
    targetCategoryHeader = GetCategoryHeaderList()
    Dim i As Integer
    Dim outputValues() As String
    Dim index As Object
    Set index = getDocumentColIndex()
    Dim count As Integer
    Dim header As Variant
    count = 0
    For i = LBound(inputValues, 1) To UBound(inputValues, 1)
        For Each header In targetCategoryHeader
            If Left(Trim(inputValues(i, index("category"))), Len(header)) = header Then
                ReDim Preserve outputValues(4, count)
                outputValues(0, count) = inputValues(i, index("category"))
                outputValues(1, count) = inputValues(i, index("itemName"))
                outputValues(2, count) = inputValues(i, index("format"))
                outputValues(3, count) = inputValues(i, index("dc"))
                outputValues(4, count) = inputValues(i, index("isr"))
                count = count + 1
            End If
        Next header
    Next i
    GetTargetValues = outputValues()
End Function
Private Function GetTargetDocument() As Workbook
    Dim editPath As New ClassEditPath
    Dim inputFolderPath As String
    Dim fileCollection As Collection
    Dim targetFilePathList() As String
    Dim targetYmd() As Long
    Dim file As Variant
    Dim fileNameBody As String
    Dim i As Integer
    Dim tempYmd As String
    Dim latestYmd As Long
    Dim targetFilePath As String
    Const documentListName = "D000 "
    Const documentYmdLength As Integer = 6
    
    i = 0
    
    inputFolderPath = editPath.GetDocumentListPath()
    Set fileCollection = GetAllFilesInFolder(inputFolderPath)
    For Each file In fileCollection
        fileNameBody = GetFileName(file)
        If Left(fileNameBody, Len(documentListName)) = documentListName Then
            ReDim Preserve targetYmd(i)
            ReDim targetFilePathList(i)
            tempYmd = Mid(fileNameBody, Len(fileNameBody) - (documentYmdLength + Len(".xlsx")) + 1, documentYmdLength)
            If IsNumeric(tempYmd) Then
                targetYmd(i) = CLng(tempYmd)
            Else
                targetYmd(i) = -1
            End If
            targetFilePathList(i) = file
            i = i + 1
        End If
    Next file
    
    latestYmd = GetMaxValue(targetYmd)
    For i = LBound(targetYmd) To UBound(targetYmd)
        If targetYmd(i) = latestYmd Then
            targetFilePath = targetFilePathList(i)
            Exit For
        End If
    Next i
    
    Set GetTargetDocument = Workbooks.Open(targetFilePath)
    
End Function
