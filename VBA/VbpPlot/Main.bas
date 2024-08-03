Attribute VB_Name = "Main"
Option Explicit

'<�T�v>
'�w�肳�ꂽ�@�z���̃t�H���_(�T�u�t�H���_�܂�)����g���q�uvbp�v�̃t�@�C����T�����A���������ꍇ��vbp�t�@�C���̃p�X��vbp�t�@�C���̓��e��V�K�V�[�g�ɏo�͂���B
'
'<�ڍ�>
'�E����FExcel VBA
'�E�v���O�����ւ̓��́F
'�@�@�t�H���_�p�X(��΃p�X)
'�E�o�̓V�[�g�̎d�l�F
'�@(a) A2�`An�Z���F��������vbp�t�@�C���̐�΃p�X
'�@(b) B1�`n1�Z���Fvbp�t�@�C���̓��e(�L�[��)
'�@(c) B2�`n2�Z���Fvbp�t�@�C���̓��e(�L�[���ɑΉ�����l)
'
'<�����>
'�Evbp�t�@�C�����ȉ��̃p�X�ɑ��݂���Ƃ���B
'�@(#1)  C:\tmp\test1.vbp
'�@(#2)  C:\tmp\sub\test2.vbp
'�E���ꂼ��vbp�t�@�C���̒��g�͈ȉ��Ƃ���B
'�@(test1.vbp)
'�@Type=Exe
'�@Form=frmMain1.frm
'�@Command32=""
'�@Name="TestProject1"
'
'�@(test2.vbp)
'�@Type=Exe
'�@Form=frmMain2.frm
'�@ExeName32="TestProject2.exe"
'�@Command32=""
'�@Name="TestProject2"
'
'�E���̏�Ԃ�Excel VBA�}�N���ɁuC:\tmp�v���w�肷��Əo�͂����V�[�g�͈ȉ��ƂȂ邱�ƁB
'
'A2�FC:\tmp\test1.vbp
'A3�FC:\tmp\sub\test2.vbp
'B1�FType
'C1�FForm
'D1�FCommand32
'E1�FName
'F1�FExeName32
'B2�FExe
'C2�FfrmMain1.frm
'D2�F""
'E2�F"TestProject1"
'F2�F""
'B3�FExe
'C3�FfrmMain2.frm
'D3�F""
'E3�F"TestProject2"
'F3�F"TestProject2.exe"
'------

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("VBP�v���b�g�����s���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("A3").value = "������..."

    Dim msg As String: msg = "����ɏI�����܂���"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VbpPlot.log"
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
    Const Clm = "O7"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(Clm).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

