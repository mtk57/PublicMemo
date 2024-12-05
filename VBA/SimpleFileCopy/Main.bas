Attribute VB_Name = "Main"
Option Explicit

Private Const CELL_IS_DEBUG = "O5"
Private Const CELL_PROCESSING = "A3"
Private Const CELL_COUNT = "A10"

Private Const STR_PROCESSING = "������..."
Private Const STR_COUNT = "Copy Success Count:"

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("�V���v���t�@�C���R�s�[�����s���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range(CELL_PROCESSING).value = STR_PROCESSING
    main_sheet.Range(CELL_COUNT).value = STR_COUNT

    Dim msg As String: msg = "����ɏI�����܂���"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "SimpleFileCopy.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "��Start"

    Worksheets("main").Activate
    Process.Run

    main_sheet.Range(CELL_COUNT).value = STR_COUNT & Process.GetResult()

    Common.WriteLog "��End"
    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description
    main_sheet.Range(CELL_PROCESSING).value = msg
    main_sheet.Range(CELL_COUNT).value = STR_COUNT

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    main_sheet.Range(CELL_PROCESSING).value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Public Sub Clear_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("�N���A���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Worksheets("main").Activate
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Process.Clear

    GoTo FINISH
    
ErrorHandler:
    MsgBox "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description

FINISH:
    Application.DisplayAlerts = True
    
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

