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
Private proc_type As PROCESS_TYPE

Private git_url As String
Private git_path As String
Private base_folder As String
Private base_branch As String
Private is_update_remote As Boolean
Private is_update_remote_s As String
Private is_delete_exist_branch As Boolean
Private is_delete_exist_branch_s As String
Private is_ignore_not_ref As Boolean
Private is_ignore_not_ref_s As String

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Git Param] "

    proc_type = UNKNOWN
    git_url = ""
    git_path = ""
    base_folder = ""
    base_branch = ""
    is_update_remote = False
    is_delete_exist_branch = False
    is_ignore_not_ref = False
End Sub

'初期化
Public Sub Init()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    Const CLM = Define.GIT_PARAM_CLM
    Const i = Define.GIT_PARAM_ROW
    
    git_url = sheet.Range(CLM & i).value
    git_path = sheet.Range(CLM & i + 1).value
    base_folder = sheet.Range(CLM & i + 2).value
    base_branch = sheet.Range(CLM & i + 3).value
    is_update_remote_s = sheet.Range(CLM & i + 4).value
    is_delete_exist_branch_s = sheet.Range(CLM & i + 5).value
    is_ignore_not_ref_s = sheet.Range(CLM & i + 6).value
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
    
    If is_update_remote_s = "NO" Or is_update_remote_s = "" Then
        is_update_remote = False
    Else
        is_update_remote = True
    End If
    
    If is_delete_exist_branch_s = "NO" Or is_delete_exist_branch_s = "" Then
        is_delete_exist_branch = False
    Else
        is_delete_exist_branch = True
    End If

    If is_ignore_not_ref_s = "NO" Or is_ignore_not_ref_s = "" Then
        is_ignore_not_ref = False
    Else
        is_ignore_not_ref = True
    End If

End Sub

'=========================================
'Getter
'=========================================

Public Function GetProcessType() As PROCESS_TYPE
    GetProcessType = proc_type
End Function

'-------------------------------------------------
'GitリポジトリURL
'-------------------------------------------------
Public Function GetUrl() As String
    GetUrl = git_url
End Function

'-------------------------------------------------
'Gitフォルダパス(絶対パス)
'-------------------------------------------------
Public Function GetGitDirPath() As String
    GetGitDirPath = git_path
End Function

'-------------------------------------------------
'起点フォルダ名
'-------------------------------------------------
Public Function GetBaseFolder() As String
    GetBaseFolder = base_folder
End Function

'-------------------------------------------------
'ベースブランチ名
'-------------------------------------------------
Public Function GetBaseBranch() As String
    GetBaseBranch = base_branch
End Function

'-------------------------------------------------
'リモートブランチを更新する
'-------------------------------------------------
Public Function IsUpdateRemote() As Boolean
    IsUpdateRemote = is_update_remote
End Function

'-------------------------------------------------
'Branchが存在した場合は削除する
'-------------------------------------------------
Public Function IsDeleteExistBranch() As Boolean
    IsDeleteExistBranch = is_delete_exist_branch
End Function

'-------------------------------------------------
'VBプロジェクト未参照ファイルを除外
'-------------------------------------------------
Public Function IsIgnoreNotRef() As Boolean
    IsIgnoreNotRef = is_ignore_not_ref
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "git_url=(" & git_url & "), " & _
          "git_path=(" & git_path & "), " & _
          "base_folder=(" & base_folder & "), " & _
          "base_branch=(" & base_branch & "), " & _
          "is_update_remote=(" & is_update_remote_s & "), " & _
          "is_delete_exist_branch=(" & is_delete_exist_branch_s & "), " & _
          "is_ignore_not_ref=(" & is_ignore_not_ref_s & ")"
          
    GetAllValue = ret
End Function

'---------------------------
' Setter
'---------------------------
Public Sub SetProcessType(ByVal type_ As PROCESS_TYPE)
    proc_type = type_
End Sub


