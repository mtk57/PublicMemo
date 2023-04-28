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

'メンバー変数
Private git_url As String
Private git_path As String
Private base_folder As String
Private base_branch As String

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Git Param] "

    git_url = ""
    git_path = ""
    base_folder = ""
    base_branch = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    Const CLM = Define.GIT_PARAM_CLM
    Const i = Define.GIT_PARAM_ROW
    
    git_url = main_sheet.Range(CLM & i).value
    git_path = main_sheet.Range(CLM & i + 1).value
    base_folder = main_sheet.Range(CLM & i + 2).value
    base_branch = main_sheet.Range(CLM & i + 3).value
End Sub

'検証
Public Sub Validate()
    If git_path = "" Then
        Err.Raise 53, , HEADER & "Gitフォルダパスが指定されていません"
    End If
    
    If base_folder = "" Then
        Err.Raise 53, , HEADER & "起点フォルダ名が指定されていません"
    End If

    If base_branch = "" Then
        Err.Raise 53, , HEADER & "ベースブランチ名が指定されていません"
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

Public Function GetBaseFolder() As String
    GetBaseFolder = base_folder
End Function

Public Function GetBaseBranch() As String
    GetBaseBranch = base_branch
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "git_url=(" & git_url & "), " & _
          "git_path=(" & git_path & "), " & _
          "base_folder=(" & base_folder & "), " & _
          "base_branch=(" & base_branch & ")"
          
    GetAllValue = ret
End Function



