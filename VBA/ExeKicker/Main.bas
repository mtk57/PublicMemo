Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "ExeKicker.log"
    
    Common.WriteLog "Start"

    process.Run

    Common.WriteLog "End"

    Common.CloseLog
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Dim errmsg As String: errmsg = "エラーが発生しました：" & Err.Description
    MsgBox errmsg, vbCritical, "エラー"
    Common.WriteLog errmsg
    Common.CloseLog
    Application.DisplayAlerts = True
End Sub


Public Sub DeleteWorkDir_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    process.DelWkDir

    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical, "エラー"
    Application.DisplayAlerts = True
End Sub


