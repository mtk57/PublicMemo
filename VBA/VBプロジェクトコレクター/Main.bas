Attribute VB_Name = "Main"
Option Explicit

Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    Process.Run

    MsgBox "�I��!"
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂����F" & Err.Description, vbCritical, "�G���["
    Application.DisplayAlerts = True
End Sub

