VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����o�[�ϐ�

'MUST
Private is_enable As String
Private ini_file_path As String

'WANT
Private is_enable_addin As Boolean
Private is_skip_comment As Boolean
Private is_exec_not_diff As Boolean

Private is_enable_addin_s As String
Private is_skip_comment_s As String
Private is_exec_not_diff_s As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    'MUST
    is_enable = "STOPPER"
    ini_file_path = ""
    
    'WANT
    is_enable_addin = False
    is_skip_comment = True
    is_exec_not_diff = False
End Sub

'������
Public Function Init(ByVal start_row_num As Integer) As String
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Dim i As Integer: i = start_row_num
    Const CLM = "N"
    
    'MUST
    is_enable = main_sheet.Range(CLM & i).value
    ini_file_path = main_sheet.Range(CLM & i + 1).value
    
    'WANT
    is_enable_addin_s = main_sheet.Range(CLM & i + 2).value
    is_skip_comment_s = main_sheet.Range(CLM & i + 3).value
    is_exec_not_diff_s = main_sheet.Range(CLM & i + 4).value
    
    Init = Validate()
End Function

'����
Private Function Validate() As String
    
    'MUST
    If is_enable = "" Or _
       is_enable = "STOPPER" Then
       is_enable = "STOPPER"
    ElseIf is_enable = "ENABLE" Then
        is_enable = "ENABLE"
    Else
        is_enable = "DISABLE"
    End If

    If is_enable = "DISABLE" Or is_enable = "STOPPER" Then
        GoTo FINISH
    End If
    
    If ini_file_path = "" Then
        Validate = "�ϊ���`Ini�t�@�C���p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFile(ini_file_path) = False Then
        Validate = "�w�肳�ꂽ�ϊ���`Ini�t�@�C�������݂��܂���"
        Exit Function
    End If

    'WANT
    If is_enable_addin_s = "" Or _
       is_enable_addin_s = "NO" Then
       is_enable_addin = False
    Else
        is_enable_addin = True
    End If

    If is_skip_comment_s = "" Or _
       is_skip_comment_s = "YES" Then
       is_skip_comment = True
    Else
        is_skip_comment = False
    End If

    If is_exec_not_diff_s = "" Or _
       is_exec_not_diff_s = "NO" Then
       is_exec_not_diff = False
    Else
        is_exec_not_diff = True
    End If

FINISH:
    Validate = ""
End Function

'=========================================
'Getter
'=========================================

'MUST

'-----------------------------------------
'�L������
' Ret : "ENABLE" or "DISABLE" or "STOPPER"
'-----------------------------------------
Public Function GetEnable() As String
    GetEnable = is_enable
End Function

'-----------------------------------------
'�ϊ���`Ini�t�@�C���p�X
'-----------------------------------------
Public Function GetIniFilePath() As String
    GetIniFilePath = ini_file_path
End Function


'WANT

'-----------------------------------------
'�A�h�C����L�������邩?
'-----------------------------------------
Public Function IsEnableAddin() As Boolean
    IsEnableAddin = is_enable_addin
End Function

'-----------------------------------------
'�R�����g�s�̓X�L�b�v���邩?
'-----------------------------------------
Public Function IsSkipComment() As Boolean
    IsSkipComment = is_skip_comment
End Function

'-----------------------------------------
'�����������Ȃ�܂Ŏ��s���邩?
'-----------------------------------------
Public Function IsExecNotDiff() As Boolean
    IsExecNotDiff = is_exec_not_diff
End Function
