Attribute VB_Name = "Main"
Option Explicit

Private Const RUN_001 = "RUN_001"
Private Const RUN_002 = "RUN_002"
Private Const RUN_003 = "RUN_003"
Private Const RUN_004 = "RUN_004"
Private Const RUN_005 = "RUN_005"
Private Const RUN_006 = "RUN_006"
Private Const DEL_BRANCH = "Delete Branch"
Private Const DEL_TAG = "Delete Tag"

Public Sub Run001_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_001 & "]を実行します") = False Then
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

    Worksheets("params").Activate
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
    Worksheets("main").Activate
    MsgBox msg
End Sub

Public Sub Run002_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_002 & "]を実行します") = False Then
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

    Worksheets("params").Activate
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
    Worksheets("main").Activate
    MsgBox msg
End Sub

Public Sub DeleteBranch_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & DEL_BRANCH & "]を実行します") = False Then
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

    Worksheets("params").Activate
    Process_Delete.Run PROCESS_TYPE.DELETE_BRANCH

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("danger_zone").Activate
    MsgBox msg
End Sub

Public Sub DeleteTag_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & DEL_TAG & "]を実行します") = False Then
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

    Worksheets("params").Activate
    Process_Delete.Run PROCESS_TYPE.DELETE_TAG

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("danger_zone").Activate
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    
    Dim is_debug_log_s As String: is_debug_log_s = sheet.Range(Define.DEBUG_LOG_CELL).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

Private Sub VisibleProcessingMessage(ByVal is_visible As Boolean)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    
    sheet.Range(Define.NOW_PROCESS).value = ""
    If is_visible = True Then
        sheet.Range(Define.NOW_PROCESS).value = "処理中..."
    End If
End Sub

Public Sub Run003_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_003 & "]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_003 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("params").Activate
    Process_003.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("main").Activate
    MsgBox msg
End Sub

Public Sub Run004_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_004 & "]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_004 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("params").Activate
    Process_004.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("main").Activate
    MsgBox msg
End Sub

Public Sub Run005_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_005 & "]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_005 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("params").Activate
    Process_005.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("main").Activate
    MsgBox msg
End Sub

Public Sub Run006_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[" & RUN_006 & "]を実行します") = False Then
        Exit Sub
    End If

    VisibleProcessingMessage True
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_006 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("params").Activate
    Process_006.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    VisibleProcessingMessage False
    Worksheets("main").Activate
    MsgBox msg
End Sub
