Attribute VB_Name = "Main"
Option Explicit

Private Const RUN_001 = "RUN_001"
Private Const RUN_002 = "RUN_002"
Private Const DEL_BRANCH = "DEL_BRANCH"
Private Const DEL_TAG = "DEL_TAG"

Public Sub Run001_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[001]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
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
    VisibleProcessingMessage False
    MsgBox msg
End Sub

Public Sub Run002_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[002]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
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
    VisibleProcessingMessage False
    MsgBox msg
End Sub

Public Sub DeleteBranch_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[Delete Branch]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & DEL_BRANCH & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_Delete.Run DELETE_ENUM.TYPE_BRANCH

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    MsgBox msg
End Sub

Public Sub DeleteTag_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[Delete Tag]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & DEL_TAG & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_Delete.Run DELETE_ENUM.TYPE_TAG

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(Define.DEBUG_LOG_CELL).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

Private Sub VisibleProcessingMessage(ByVal is_visible As Boolean)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    main_sheet.Range(Define.NOW_PROCESS).value = ""
    If is_visible = True Then
        main_sheet.Range(Define.NOW_PROCESS).value = "処理中..."
    End If
End Sub

