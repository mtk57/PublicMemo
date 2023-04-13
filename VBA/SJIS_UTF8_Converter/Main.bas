Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "ExeKicker.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Process.Run

    Common.WriteLog "★End"

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

    Process.DelWkDir

    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical, "エラー"
    Application.DisplayAlerts = True
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "N"
    Const i = 18
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(CLM & i).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

