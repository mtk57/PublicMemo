Attribute VB_Name = "Main"
Option Explicit

Const LOG_NAME = "VbFileListCreator.log"

'--------------------------------------------------------
'���̃u�b�N����Ăяo���ꍇ�͂��̃��\�b�h���g������
' search_dir_path : I : ��������t�H���_�p�X (��΃p�X)
' target_prj : I : �Ώۃv���W�F�N�g(vbp/vbproj)
' ignore_files : I : ���W���O�t�@�C�� (vbproj�̂ݑΏہB���p�J���}�Ōq����)
' target_exts : I : ���W�Ώۊg���q(���p�J���}�Ōq����BEx "vb,frm,bas,cls,ctl")
' is_debug : I : �f�o�b�O���O�o�͗L��(True=�o�͂���)
' Ret : ���W���� (Dict)
'--------------------------------------------------------
Public Function Run( _
    ByVal search_dir_path As String, _
    ByVal target_prj As String, _
    ByVal ignore_files As String, _
    ByVal target_exts As String, _
    ByVal is_debug As Boolean _
) As Dict
    
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Process.IS_EXTERNAL = True
    
    Dim msg As String: msg = "����ɏI�����܂���"
    Dim ret As Dict
    
    If is_debug = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + LOG_NAME
    End If

    '�J�n
    Common.WriteLog "------------------------------------"
    Common.WriteLog "��Start"
    Common.WriteLog "Receive Param=(" & search_dir_path & "), " & _
                                  "(" & target_prj & "), " & _
                                  "(" & ignore_files & "), " & _
                                  "(" & target_exts & "), " & _
                                  "(" & is_debug & ")"

    CreateParamForExternal search_dir_path, target_prj, ignore_files, target_exts
    Set ret = Process.RunForExternal

    Common.WriteLog "��End"
    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description
    Set ret = Empty

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    Set Run = ret
End Function

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
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + LOG_NAME
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

Private Sub CreateParamForExternal( _
    ByVal search_dir_path As String, _
    ByVal target_prj As String, _
    ByVal ignore_files As String, _
    ByVal target_exts As String _
)
    Common.WriteLog "CreateParamForExternal S"
    
    Dim main_param As MainParam
    Set main_param = New MainParam

    main_param.InitForExternal search_dir_path, target_prj, ignore_files, target_exts
    
    Set Process.main_param = main_param
    
    Common.WriteLog "CreateParamForExternal E"
End Sub

