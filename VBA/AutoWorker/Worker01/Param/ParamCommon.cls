VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'�����o�[�ϐ�
Private external_path As String
Private dst_dir_path As String
Private is_debug_log As Boolean
Private is_debug_log_s As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    HEADER = "[Common Param] "

    external_path = ""
    dst_dir_path = ""
    is_debug_log = False
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    Const CLM = Define.COMMON_PARAM_CLM
    Const i = Define.COMMON_PARAM_ROW
    
    external_path = main_sheet.Range(CLM & i).value
    dst_dir_path = main_sheet.Range(CLM & i + 1).value
    is_debug_log_s = main_sheet.Range(CLM & i + 2).value
End Sub

'����
Public Sub Validate()
    If external_path = "" Then
        Err.Raise 53, , HEADER & "�O���c�[���t�@�C���p�X���w�肳��Ă��܂���"
    End If
    
    If Common.IsExistsFile(external_path) = False Then
        Err.Raise 53, , HEADER & "�O���c�[���t�@�C���p�X��������܂��� (" & external_path & ")"
    End If
    
    If dst_dir_path = "" Then
        Err.Raise 53, , HEADER & "�R�s�[��t�H���_�p�X���w�肳��Ă��܂���"
    End If
    
    If is_debug_log_s = "" Or is_debug_log_s = "NO" Then
        is_debug_log = False
    Else
        is_debug_log = True
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetExternalPath() As String
    GetExternalPath = external_path
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

Public Function IsDebugLog() As Boolean
    IsDebugLog = is_debug_log
End Function


Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "external_path=(" & external_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_debug_log=(" & is_debug_log_s & ")"
          
    GetAllValue = ret
End Function
