Attribute VB_Name = "Main"
Option Explicit

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("�Z�����������F�ύX�����s���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim MAIN_SHEET As Worksheet
    Set MAIN_SHEET = ThisWorkbook.Sheets("main")
    MAIN_SHEET.Range("A3").value = "������..."

    Dim msg As String: msg = "����ɏI�����܂���"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "�Z�����������F�ύX.log"
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


