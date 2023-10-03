Attribute VB_Name = "ConstantsModule"
' Tools > Reference Settings > Microsoft Word 16.0 Object Library
' Tools > Reference Settings > Microsoft Scripting runtime
Option Explicit

Public Const debugFlag As Boolean = True
Public Const isr As String = "情報システム研究室"

Public Sub main()
    Call ExecConvertToPdfLatestFile
    Call ExecCreateTextFunction
    MsgBox "処理が終了しました"
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
