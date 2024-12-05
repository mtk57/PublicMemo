Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("With句解析を実行します") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("A3").value = "処理中..."

    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "WithParser.log"
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
    main_sheet.Range("A3").value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Public Sub Clear_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("クリアします") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Worksheets("main").Activate
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Process.Clear

    GoTo FINISH
    
ErrorHandler:
    MsgBox "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description

FINISH:
    Application.DisplayAlerts = True
    
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const Clm = "L6"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(Clm).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

