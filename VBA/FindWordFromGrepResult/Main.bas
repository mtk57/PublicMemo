Attribute VB_Name = "Main"
Option Explicit

Private Const CELL_IS_DEBUG = "O12"
Private Const CELL_PROCESSING = "A3"

Private Const STR_PROCESSING = "処理中..."

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("FindWordFromGrepResultを実行します") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range(CELL_PROCESSING).value = STR_PROCESSING

    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "FindWordFromGrepResult.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("main").Activate
    Process.Run

    Common.WriteLog "★End"
    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description
    main_sheet.Range(CELL_PROCESSING).value = msg

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    main_sheet.Range(CELL_PROCESSING).value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(CELL_IS_DEBUG).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function



