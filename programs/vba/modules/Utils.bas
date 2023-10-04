Attribute VB_Name = "Utils"
Option Explicit
Public Function GetFiscalYear() As Integer
    Dim currentDate As Date
    Dim fiscalYear As Integer
    
    ' Get the current date
    currentDate = Date
    
    ' Compare with the start date of the fiscal year (April 1st) to determine the fiscal year
    If currentDate >= DateSerial(Year(currentDate), 4, 1) Then
        fiscalYear = Year(currentDate)
    Else
        fiscalYear = Year(currentDate) - 1
    End If
    
    ' Return the result
    GetFiscalYear = fiscalYear
End Function

Public Function ToNFC(str) As String

    Dim i As Integer
    Dim nfdChrs As String
    Dim nfcChr As String
    i = 1
    
    Do
        i = InStr(str, ChrW(12443))             '濁点
        If i = 0 Then Exit Do
        nfdChrs = Mid(str, i - 1, 2)
        nfcChr = ChrW(AscW(Mid(str, i - 1, 1)) + 1)
        str = Replace(str, nfdChrs, nfcChr)
    Loop
    Do
        i = InStr(str, ChrW(12444))             '半濁点
        If i = 0 Then Exit Do
        nfdChrs = Mid(str, i - 1, 2)
        nfcChr = ChrW(AscW(Mid(str, i - 1, 1)) + 2)
        str = Replace(str, nfdChrs, nfcChr)
    Loop

    ToNFC = str

End Function



