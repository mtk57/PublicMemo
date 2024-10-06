Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("セル差分文字色変更を実行します") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim MAIN_SHEET As Worksheet
    Set MAIN_SHEET = ThisWorkbook.Sheets("main")
    MAIN_SHEET.Range("A3").value = "処理中..."

    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "セル差分文字色変更.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("main").Activate
    Process.Run

    Common.WriteLog "★End"
    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    MAIN_SHEET.Range("A3").value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim MAIN_SHEET As Worksheet
    Set MAIN_SHEET = ThisWorkbook.Sheets("main")
    Const Clm = "F5"
    
    Dim is_debug_log_s As String: is_debug_log_s = MAIN_SHEET.Range(Clm).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function


