Attribute VB_Name = "CreateText"
Option Explicit


Public Sub testcreatetext()
    Dim targetWorksheet As Variant
    Call GetTargetWorksheets
End Sub
Private Sub GetTargetWorksheets()
    Dim wb As Workbook
    Dim i As Integer
    Dim targetSheetName(1) As String
    targetSheetName(0) = "文書管理台帳(2)"
    targetSheetName(1) = "文書管理台帳(3)"
    Dim targetCategoryHeader(1) As String
    targetCategoryHeader(0) = "QF"
    targetCategoryHeader(1) = "ISF"
    Dim targetSheetValues() As Variant
    ReDim targetSheetValues(UBound(targetSheetName))
    Set wb = GetTargetDocument()
    If wb Is Nothing Then
        Exit Sub
    End If
    For i = LBound(targetSheetName) To UBound(targetSheetName)
        targetSheetValues(i) = wb.Worksheets(targetSheetName(i)).UsedRange.Value
    Next i
    wb.Close savechanges:=False
    Dim targetValuesList() As String
    Dim targetValues() As String
    Dim dummy As Variant
    For i = LBound(targetSheetValues) To UBound(targetSheetValues)
        targetValues = GetTargetValues(targetSheetValues(i), targetCategoryHeader(i))
        dummy = CreateFileNames(targetValues)
        
    Next i
Debug.Print 0
End Sub
Private Function CreateFileNames(targetValues() As String) As Variant
    Dim index As Object
    Dim dummy As String
    Set index = CreateTargetValuesIndex()
    Dim i As Integer
    Dim arr() As String
    ReDim arr(UBound(targetValues, 1))
    For i = LBound(targetValues, 2) To UBound(targetValues, 2)
        Dim j As Integer
        For j = LBound(targetValues, 1) To UBound(targetValues, 1)
            arr(j) = targetValues(j, i)
        Next j
        dummy = CreateFileName(arr, index)
        Debug.Print dummy
    Next i
    CreateFileNames = ""
End Function
Private Function CreateFileName(values() As String, index As Object) As String
    Const constRef As String = "参照"
    Dim textHeader As String
    Dim temp As Variant
    Dim refCategory As String
    Dim targetDept As String
    Const dc As String = "データ管理室"
    Const isr As String = "情報システム研究室"
    If values(index("isr")) = "○" Then
        CreateFileName = ""
        Exit Function
    End If
    ' 「【区分】【記録名】」までは共通
    textHeader = values(index("category")) & " " & values(index("itemName"))
    Dim fileName As String
    fileName = textHeader
    ' ISRに…参照の文字列が存在する場合、ファイル名は「【区分】【記録名】情報システム研究室【区分】参照」
    If InStr(1, values(index("isr")), constRef) > 0 Then
        targetDept = isr
        refCategory = GetRefCategoryName(values(index("isr")), constRef)
        fileName = "*"
    ' DCに…参照の文字列が存在する場合、ファイル名は「【区分】【記録名】データ管理室【区分】参照」
    ElseIf InStr(1, values(index("dc")), constRef) > 0 Then
        targetDept = dc
        refCategory = GetRefCategoryName(values(index("dc")), constRef)
        fileName = "#"
    ' ISRが空白でDCが○の場合、ファイル名は「【区分】【記録名】データ管理室【区分】参照」
    ElseIf values(index("isr")) = "" And values(index("dc")) <> "" Then
        targetDept = dc
        refCategory = ""
        fileName = fileName & " " & targetDept & values(index("category")) & constRef & ".txt"
    End If
    ' 参照先の区分を取得する
    
    CreateFileName = fileName

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
Private Function GetTargetValues(inputValues, header) As String()
    Dim i As Integer
    Dim outputValues() As String
    Dim index As Object
    Set index = getDocumentColIndex()
    Dim count As Integer
    count = 0
    For i = LBound(inputValues, 1) To UBound(inputValues, 1)
        If Left(Trim(inputValues(i, index("category"))), Len(header)) = header Then
            ReDim Preserve outputValues(4, count)
            outputValues(0, count) = inputValues(i, index("category"))
            outputValues(1, count) = inputValues(i, index("itemName"))
            outputValues(2, count) = inputValues(i, index("format"))
            outputValues(3, count) = inputValues(i, index("dc"))
            outputValues(4, count) = inputValues(i, index("isr"))
            count = count + 1
        End If
    Next i
    GetTargetValues = outputValues()
End Function
Private Function GetTargetDocument() As Workbook
    Dim editPath As New ClassEditPath
    Dim inputFolderPath As String
    Dim fileCollection As collection
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


Private Sub GenerateISOAuditFiles()
    ' Clear temporary variables
    Dim kInputPath As String
    Dim fiscal_year As String
    Dim kOutputParentPath As String
    Dim kQmfHeader As String
    Dim kIsmsHeader As String
    Dim kIsms As String
    Dim kQms As String
    Dim kRefIsms As String
    Dim kRefQms As String
    Dim kRefCommon As String
    Dim kTargetSign As String
    Dim kIsrColName As String
    Dim kDcColName As String
    Dim kCategory As String
    Dim kFormat As String
    Dim kItemName As String
    Dim kIsrName As String
    Dim kDcName As String
    Dim kPaper As String
    
    ' Assign constants
    kInputPath = "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/02 文書（ト゛ラフト）/D000 QMS・ISMS文書一覧 230922.xlsx"
    fiscal_year = CalculateTodayFiscalYear()
    kOutputParentPath = "~/Library/CloudStorage/Box-Box/Projects/ISO/QMS・ISMS文書/04 記録/" & fiscal_year & "年度/"
    kQmfHeader = "QF"
    kIsmsHeader = "ISF"
    kIsms = "ISMS"
    kQms = "QMS"
    kRefIsms = kIsms & "参照"
    kRefQms = kQms & "参照"
    kRefCommon = "共通"
    kTargetSign = "○"
    kIsrColName = "...10"
    kDcColName = "保管部門"
    kCategory = "区分"
    kFormat = "形式"
    kItemName = "記録名"
    kIsrName = "情報システム研究室"
    kDcName = "データ管理室"
    kPaper = "紙"
    
    ' Lock constants
    ' (No equivalent VBA code is needed)
    
    ' Input file path
    Dim input_path As String
    input_path = kInputPath
    
    ' Read all sheets from the Excel file
    Dim wb As Workbook
    Dim all_sheets As Sheets
    Dim sheet As Worksheet
    Set wb = Workbooks.Open(fileName:=input_path)
    Set all_sheets = wb.Sheets
    Dim sheet_data As collection
    Set sheet_data = New collection
    
    For Each sheet In all_sheets
        ' Assuming you want to skip the first 4 rows when reading each sheet
        Dim data As ListObject
        Set data = sheet.ListObjects.Add(xlSrcRange, sheet.Range("A5").CurrentRegion, , xlYes)
        sheet_data.Add data.DataBodyRange.Value
    Next sheet
    
    ' Close the input workbook
    wb.Close savechanges:=False
    
    ' Main processing code goes here...
    ' You'll need to translate the R code to VBA.
    ' It may involve creating subroutines for functions like FilterTargetDs, GenerateFilenames, etc.
    ' and then calling those subroutines in the appropriate order.
    
    ' For example, you can create a subroutine for FilterTargetDs like this:
    ' Sub FilterTargetDs(ds As Range)
    '    ' Translate the R code for FilterTargetDs to VBA here...
    ' End Sub
    
    ' Once you've translated the R code to VBA, call the subroutines to execute the main processing.
    
    ' Output files
    ' You'll need to add VBA code to write output files.
    
End Sub

Function CalculateTodayFiscalYear() As Integer
    ' Translate the R code for CalculateTodayFiscalYear to VBA here...
    ' Return the fiscal year as an integer
End Function

' Define other subroutines and functions as needed...


