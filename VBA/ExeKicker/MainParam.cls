VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����o�[�ϐ�

'MUST
Private exe_file_path As String
Private src_dir_path As String
Private dst_dir_path As String
Private in_ext As String
Private addin_file_path As String
Private exe_ini_file_path As String
Private winmerge_file_path As String

'WANT
Private is_contain_sub_dir As Boolean
Private ignore_files As String
Private tool_wk_dir_path As String
Private is_step_wk_dir As Boolean
Private is_del_wk_dir As Boolean
Private max_exec_cnt As Integer

Private is_contain_sub_dir_s As String
Private is_step_wk_dir_s As String
Private is_del_wk_dir_s As String
Private max_exec_cnt_s As String

Private ignore_file_path As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    'MUST
    exe_file_path = ""
    src_dir_path = ""
    dst_dir_path = ""
    in_ext = ""
    addin_file_path = ""
    exe_ini_file_path = ""
    winmerge_file_path = ""
    
    'WANT
    is_contain_sub_dir = False
    ignore_files = ""
    tool_wk_dir_path = ""
    is_step_wk_dir = False
    is_del_wk_dir = True
    max_exec_cnt = 0
End Sub

'������
Public Function Init() As String
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "N"
    Const i = 6
    
    'MUST
    exe_file_path = main_sheet.Range(CLM & i).value
    src_dir_path = main_sheet.Range(CLM & i + 1).value
    dst_dir_path = main_sheet.Range(CLM & i + 2).value
    in_ext = main_sheet.Range(CLM & i + 3).value
    addin_file_path = main_sheet.Range(CLM & i + 4).value
    exe_ini_file_path = main_sheet.Range(CLM & i + 5).value
    winmerge_file_path = main_sheet.Range(CLM & i + 6).value
    
    'WANT
    is_contain_sub_dir_s = main_sheet.Range(CLM & i + 7).value
    ignore_files = main_sheet.Range(CLM & i + 8).value
    tool_wk_dir_path = main_sheet.Range(CLM & i + 9).value
    is_step_wk_dir_s = main_sheet.Range(CLM & i + 10).value
    is_del_wk_dir_s = main_sheet.Range(CLM & i + 11).value
    max_exec_cnt_s = main_sheet.Range(CLM & i + 12).value
    
    Init = Validate()
End Function

'����
Private Function Validate() As String
    'MUST
    If exe_file_path = "" Then
        Validate = "exe�t�@�C���p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFile(exe_file_path) = False Then
        Validate = "�w�肳�ꂽexe�t�@�C�������݂��܂���"
        Exit Function
    End If

    If src_dir_path = "" Then
        Validate = "SRC�t�H���_�p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Validate = "�w�肳�ꂽSRC�t�H���_�����݂��܂���"
        Exit Function
    End If

    If dst_dir_path = "" Then
        Validate = "DST�t�H���_�p�X���w�肳��Ă��܂���"
        Exit Function
    End If

    If in_ext = "" Then
        Validate = "�g���q���w�肳��Ă��܂���"
        Exit Function
    End If

    If addin_file_path = "" Then
        Validate = "�A�h�C���t�@�C���p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFile(addin_file_path) = False Then
        Validate = "�w�肳�ꂽ�A�h�C���t�@�C�������݂��܂���"
        Exit Function
    End If

    If exe_ini_file_path = "" Then
        Validate = "ExeIni�t�@�C���p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFile(exe_ini_file_path) = False Then
        Validate = "�w�肳�ꂽExeIni�t�@�C�������݂��܂���"
        Exit Function
    End If

    If winmerge_file_path = "" Then
        Validate = "WinMerge�t�@�C���p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFile(winmerge_file_path) = False Then
        Validate = "�w�肳�ꂽWinMerge�t�@�C�������݂��܂���"
        Exit Function
    End If

    'WANT
    If is_contain_sub_dir_s = "" Or _
       is_contain_sub_dir_s = "NO" Then
       is_contain_sub_dir = False
    Else
        is_contain_sub_dir = True
    End If

    If is_step_wk_dir_s = "" Or _
       is_step_wk_dir_s = "NO" Then
       is_step_wk_dir = False
    Else
        is_step_wk_dir = True
    End If

    If is_del_wk_dir_s = "" Or _
       is_del_wk_dir_s = "YES" Then
       is_del_wk_dir = True
    Else
        is_del_wk_dir = False
    End If

    If max_exec_cnt_s = "" Then
        max_exec_cnt = 1
    ElseIf IsNumeric(max_exec_cnt) Then
        max_exec_cnt = Int(max_exec_cnt_s)
        
        If max_exec_cnt > 10 Then
            max_exec_cnt = 10
        ElseIf max_exec_cnt < 1 Then
            max_exec_cnt = 1
        End If
    Else
        Validate = "�ő���s�񐔂������ł͂���܂���"
        Exit Function
    End If

    Validate = ""
End Function

'=========================================
'Getter
'=========================================

'MUST

'-----------------------------------------
'exe�t�@�C���p�X
'-----------------------------------------
Public Function GetExeFilePath() As String
    GetExeFilePath = exe_file_path
End Function

'-----------------------------------------
'SRC�t�H���_�p�X
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

'-----------------------------------------
'DST�t�H���_�p�X
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDstDirPath = dst_dir_path
End Function

'-----------------------------------------
'�g���q
'-----------------------------------------
Public Function GetInExtension() As String
    GetInExtension = in_ext
End Function

'-----------------------------------------
'�A�h�C���t�@�C���p�X
'-----------------------------------------
Public Function GetAddinFilePath() As String
    GetAddinFilePath = addin_file_path
End Function

'-----------------------------------------
'ExeIni�t�@�C���p�X
'-----------------------------------------
Public Function GetExeIniFilePath() As String
    GetExeIniFilePath = exe_ini_file_path
End Function

'-----------------------------------------
'WinMerge�t�@�C���p�X
'-----------------------------------------
Public Function GetWinMergeFilePath() As String
    GetWinMergeFilePath = winmerge_file_path
End Function


'WANT

'-----------------------------------------
'�T�u�t�H���_�܂ނ�?
'-----------------------------------------
Public Function IsContainSubDir() As Boolean
    IsContainSubDir = is_contain_sub_dir
End Function

'-----------------------------------------
'���O�t�@�C�����X�g
'-----------------------------------------
Public Function GetIgnoreFiles() As String()
    GetIgnoreFiles = Split(ignore_files, ",")
End Function

Public Function GetIgnoreFilePath() As String
    GetIgnoreFilePath = ignore_file_path
End Function

'-----------------------------------------
'�c�[����Ɨp�t�H���_�p�X
'-----------------------------------------
Public Function GetToolWorkDirPath() As String
    GetToolWorkDirPath = tool_wk_dir_path
End Function

'-----------------------------------------
'��Ɨp�t�H���_�ɓr���o�߂��c����?
'-----------------------------------------
Public Function IsStepWorkDir() As Boolean
    IsStepWorkDir = is_step_wk_dir
End Function

'-----------------------------------------
'���s��ɍ�Ɨp�t�H���_�p�X���폜���邩?
'-----------------------------------------
Public Function IsDeleteWorkDir() As Boolean
    IsDeleteWorkDir = is_del_wk_dir
End Function

'-----------------------------------------
'�ő���s��
'-----------------------------------------
Public Function GetMaxExecCount() As Integer
    GetMaxExecCount = max_exec_cnt
End Function


'=========================================
'Setter
'=========================================

'-----------------------------------------
'���O�t�@�C�����X�g
'-----------------------------------------
Public Sub SetIgnoreFilePath(ByVal path As String)
    ignore_file_path = path
End Sub

'-----------------------------------------
'�c�[����Ɨp�t�H���_�p�X
'-----------------------------------------
Public Sub SetToolWorkDirPath(ByVal path As String)
    tool_wk_dir_path = path
End Sub