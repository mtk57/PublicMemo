Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    process.Run

    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂����F" & Err.Description, vbCritical, "�G���["
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


