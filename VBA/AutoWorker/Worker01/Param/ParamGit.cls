VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamGit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'�����o�[�ϐ�
Private git_url As String
Private git_path As String
Private base_branch As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    HEADER = "[Git Param] "

    git_url = ""
    git_path = ""
    base_branch = ""
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    Const CLM = Define.GIT_PARAM_CLM
    Const i = Define.GIT_PARAM_ROW
    
    git_url = main_sheet.Range(CLM & i).value
    git_path = main_sheet.Range(CLM & i + 1).value
    base_branch = main_sheet.Range(CLM & i + 2).value
End Sub

'����
Public Sub Validate()
    If git_url = "" Then
        Err.Raise 53, , HEADER & "Git���|�W�g��URL���w�肳��Ă��܂���"
    End If
    
    If git_path = "" Then
        Err.Raise 53, , HEADER & "Git�t�H���_�p�X���w�肳��Ă��܂���"
    End If
    
    If Common.IsExistsFolder(git_path) = False Then
        Err.Raise 53, , HEADER & "Git�t�H���_�p�X��������܂��� (" & git_path & ")"
    End If

    If base_branch = "" Then
        Err.Raise 53, , HEADER & "�x�[�X�u�����`�����w�肳��Ă��܂���"
    End If

End Sub

'=========================================
'Getter
'=========================================

Public Function GetUrl() As String
    GetUrl = git_url
End Function

Public Function GetGitDirPath() As String
    GetGitDirPath = git_path
End Function

Public Function GetBaseBranch() As String
    GetBaseBranch = base_branch
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "git_url=(" & git_url & "), " & _
          "git_path=(" & git_path & "), " & _
          "base_branch=(" & base_branch & ")"
          
    GetAllValue = ret
End Function


