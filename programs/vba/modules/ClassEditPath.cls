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
    fiscalYear = GetFiscalYear() - 1
    Dim parentPathHeaderHead As String
    parentPathHeaderHead = "C:\Users\Mariko\Box\Projects\ISO\QMS�EISMS����\04 �L�^\"
    parentPathHeader = parentPathHeaderHead & fiscalYear & "�N�x\"
End Sub

Public Function GetInputPath(folderName As String) As String
    GetInputPath = parentPathHeader & "�h���t�g\" & folderName & "\"
End Function

Public Function GetOutputPath(folderName As String) As String
    If debugFlag Then
        GetOutputPath = "C:\Users\Mariko\Box\Datacenter\Users\ohtsuka\2023\20230926\output\" & folderName & "\"
    Else
        GetOutputPath = parentPathHeader & "�Œ�\" & folderName & "\"
    End If
End Function
