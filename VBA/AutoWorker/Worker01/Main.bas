Attribute VB_Name = "Main"
Option Explicit

Private Const RUN_001 = "RUN_001"
Private Const RUN_002 = "RUN_002"
Private Const DEBUG_LOG_CLM = "D14"

Public Sub Run001_Click()
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_001 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_001.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Public Sub Run002_Click()
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_002 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_002.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(DEBUG_LOG_CLM).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

