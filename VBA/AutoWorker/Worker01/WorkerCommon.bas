Attribute VB_Name = "WorkerCommon"
Option Explicit

'ローカル/リモートブランチの存在チェック
Public Function IsExistBranch(ByVal repo_path As String, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistBranch S"
    
    Dim is_exist_local As Boolean: is_exist_local = IsExistLocalBranch(repo_path, branch)
    Dim is_exist_remote As Boolean: is_exist_remote = IsExistRemoteBranch(repo_path, branch)
    
    IsExistBranch = False
    If is_exist_local = True And is_exist_remote = True Then
         IsExistBranch = True
    End If
    
    Common.WriteLog "IsExistBranch E"
End Function

'ローカルブランチの存在チェック
Public Function IsExistLocalBranch(ByVal repo_path As String, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistLocalBranch S"
    
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
Public Function IsExistRemoteBranch(ByVal repo_path As String, ByVal branch As String) As Boolean
    Common.WriteLog "IsExistRemoteBranch S"
    
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
Public Function IsExistTag(ByVal repo_path As String, ByVal tag As String) As Boolean
    Common.WriteLog "IsExistTag S"
    
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
