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
    Dim errmsg As String: errmsg = "�G���[���������܂����F" & Err.Description
    MsgBox errmsg, vbCritical, "�G���["
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
    MsgBox "�G���[���������܂����F" & Err.Description, vbCritical, "�G���["
    Application.DisplayAlerts = True
End Sub


