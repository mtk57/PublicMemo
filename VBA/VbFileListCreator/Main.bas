Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("VB�t�@�C�����X�g�̍쐬�����s���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("A3").value = "������..."

    Dim msg As String: msg = "����ɏI�����܂���"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VBPrjCollector.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "��Start"

    Worksheets("main").Activate
    Process.Run

    Common.WriteLog "��End"
    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    main_sheet.Range("A3").value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const clm = "O7"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(clm).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function
