Attribute VB_Name = "Main"
Option Explicit

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CELL = "D10"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(CELL).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

Public Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "SJIS_UTF8_Converter.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "��Start"

    Process.Run

    Common.WriteLog "��End"

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

    Process.DelWkDir

    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂����F" & Err.Description, vbCritical, "�G���["
    Application.DisplayAlerts = True
End Sub



