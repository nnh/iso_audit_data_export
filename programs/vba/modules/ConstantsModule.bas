Attribute VB_Name = "ConstantsModule"
' Tools > Reference Settings > Microsoft Word 16.0 Object Library
' Tools > Reference Settings > Microsoft Scripting runtime
Option Explicit

Public Const debugFlag As Boolean = True
Public Const isr As String = "���V�X�e��������"

Public Sub main()
    Call ExecConvertToPdfLatestFile
    Call ExecCreateTextFunction
    MsgBox "�������I�����܂���"
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
