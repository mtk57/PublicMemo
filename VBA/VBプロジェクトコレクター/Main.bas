Attribute VB_Name = "Main"
Option Explicit

Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    Process.Run

    MsgBox "終了!"
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical, "エラー"
    Application.DisplayAlerts = True
End Sub

