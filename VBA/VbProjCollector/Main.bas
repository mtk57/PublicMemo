Attribute VB_Name = "Main"
Option Explicit

'--------------------------------------------------------
'���̃u�b�N����Ăяo���ꍇ�͂��̃��\�b�h���g������
' vbprj_files : I : VB�v���W�F�N�g�t�@�C���p�X���X�g(��΃p�X)
' dst_dir_path : I : �R�s�[��t�H���_�p�X(��΃p�X)
' base_folder : I : �ړ��N�_�t�H���_��
' is_create_build_bat : I : �r���hBAT�o�͗L��(True=�o�͂���)
' is_debug : I : �f�o�b�O���O�o�͗L��(True=�o�͂���)
' Ret : True/False (True=����)
'--------------------------------------------------------
Public Function Run( _
    ByRef vbprj_files() As String, _
    ByVal dst_dir_path As String, _
    ByVal base_folder As String, _
    ByVal is_create_build_bat As Boolean, _
    ByVal is_debug As Boolean _
    ) As Boolean
    
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Process.IS_EXTERNAL = True
    
    Dim msg As String: msg = "����ɏI�����܂���"
    Dim ret As Boolean: ret = True
    
    If is_debug = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VBPrjCollector.log"
    End If

    '�J�n
    Common.WriteLog "------------------------------------"
    Common.WriteLog "��Start"
    Common.WriteLog "Receive Param=(" & dst_dir_path & "), (" & base_folder & ")"
    
    CreateParamForExternal vbprj_files, dst_dir_path, base_folder, is_create_build_bat
    Process.Run

    Common.WriteLog "��End"
    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description
    ret = False

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    Run = ret
End Function

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("VB�v���W�F�N�g�t�@�C�����W�����s���܂�") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("A3").value = "������..."
    
    Process.IS_EXTERNAL = False

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
    Const CLM = "O8"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(CLM).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

Private Sub CreateParamForExternal( _
    ByRef vbprj_files() As String, _
    ByVal dst_dir_path As String, _
    ByVal base_dir As String, _
    ByVal is_create_build_bat As Boolean _
    )
    Common.WriteLog "CreateParamForExternal S"
    
    Dim main_param As MainParam
    Dim sub_param As SubParam
    Set main_param = New MainParam
    Set sub_param = New SubParam
    
    main_param.InitForExternal dst_dir_path, base_dir, is_create_build_bat
    sub_param.InitForExternal vbprj_files
    
    Set Process.main_param = main_param
    Set Process.sub_param = sub_param
    
    Common.WriteLog "CreateParamForExternal E"
End Sub
