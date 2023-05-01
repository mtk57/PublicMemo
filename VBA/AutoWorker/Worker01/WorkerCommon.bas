Attribute VB_Name = "WorkerCommon"
Option Explicit

'ローカル/リモートブランチの存在チェック
Public Function IsExistBranch(ByRef prms As ParamContainer, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    Dim is_exist_local As Boolean: is_exist_local = IsExistLocalBranch(prms, branch)
    Dim is_exist_remote As Boolean: is_exist_remote = IsExistRemoteBranch(prms, branch)
    
    IsExistBranch = False
    If is_exist_local = True Or is_exist_remote = True Then
         IsExistBranch = True
    End If
    
    Common.WriteLog "IsExistBranch E"
End Function

'ローカルブランチの存在チェック
Public Function IsExistLocalBranch(ByRef prms As ParamContainer, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistLocalBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git branch --list " & branch
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistLocalBranch = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistLocalBranch = True
    End If
    
    Common.WriteLog "IsExistLocalBranch E"
End Function

'リモートブランチの存在チェック
Public Function IsExistRemoteBranch(ByRef prms As ParamContainer, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistRemoteBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git branch --list --remote origin/" & branch
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistRemoteBranch = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistRemoteBranch = True
    End If
    
    Common.WriteLog "IsExistRemoteBranch E"
End Function

'タグの存在チェック
Public Function IsExistTag(ByRef prms As ParamContainer, ByVal tag As String) As Boolean
    Common.WriteLog "IsExistTag S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git tag --list " & tag
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistTag = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistTag = True
    End If
    
    Common.WriteLog "IsExistTag E"
End Function

'VBプロジェクト名を返す
Public Function GetProjectName(ByVal vbprj_file_path As String) As String
    Common.WriteLog "GetProjectName S"
    Dim vbprj_file_name As String: vbprj_file_name = Common.GetFileName(vbprj_file_path)
    Dim ext As String: ext = Common.GetFileExtension(vbprj_file_name)
    GetProjectName = Replace(vbprj_file_name, "." & ext, "")
    Common.WriteLog "GetProjectName E"
End Function

Public Sub DoClone(ByRef prms As ParamContainer)
    Common.WriteLog "DoClone S"
    
    If prms.GetUrl() = "" Then
        Common.WriteLog "DoClone E1"
        Exit Sub
    End If
    
    If Common.IsExistsFolder(prms.GetGitDirPath()) = True Then
        Common.WriteLog "DoClone E2"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    'クローン
    cmd = "git clone " & prms.GetUrl() & " " & prms.GetGitDirPath()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoClone E"
End Sub

Public Sub SwitchDevelopBranch(ByRef prms As ParamContainer)
    Common.WriteLog "SwitchDevelopBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'developブランチが存在しない場合はエラーとする
    If IsExistBranch(prms, prms.GetBaseBranch()) = False Then
        Err.Raise 53, , "ブランチが存在しません。(" & prms.GetBaseBranch() & ")"
    End If
    
    'カレントブランチを確認する
    cmd = "git branch --show-current"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If git_result(0) = prms.GetBaseBranch() Then
        Common.WriteLog "CheckoutBranch E1"
        Exit Sub
    End If
    
    'ローカルブランチがあるか確認する
    Dim is_exist_local As Boolean: is_exist_local = False
    
    cmd = "git branch --list " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If Common.IsEmptyArray(git_result) = False Then
        is_exist_local = True
    End If
    
    If is_exist_local = True Then
        'ローカルブランチがあるのでswitchで切り替え
        cmd = "git switch " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    Else
        'ローカルブランチがないので作成して切り替え
        cmd = "git checkout -b " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If
    
    Common.WriteLog "SwitchDevelopBranch E"
End Sub

Public Sub DoPull(ByRef prms As ParamContainer)
    Common.WriteLog "DoPull S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'プル
    cmd = "git pull origin " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPull E"
End Sub


