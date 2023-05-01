Attribute VB_Name = "Process_Delete"
Option Explicit

Private prms As ParamContainer
Private SEP As String
Private DQ As String

Private Const FOR_TEST = True

Public Sub Run(ByVal delete_type As DELETE_ENUM)
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Dim msg As String: msg = ""

    Set prms = New ParamContainer
    
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    Dim i As Integer
    Dim target As ParamTarget
    Dim targetlist() As ParamTarget
    targetlist = prms.GetTargetList()
    
    For i = LBound(targetlist) To UBound(targetlist)
    
        Set target = targetlist(i)
    
        If delete_type = TYPE_BRANCH Then
            DeleteBranch target
        ElseIf delete_type = TYPE_TAG Then
            DeleteTag target
        End If
        
    Next i
        
    Common.WriteLog "Run E"
End Sub

Private Sub DeleteBranch(ByRef target As ParamTarget)
    Common.WriteLog "DeleteBranch S"
    
    If FOR_TEST = False Then
        Common.WriteLog "DeleteBranch E1"
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("[For TEST] ブランチの削除をを実行します。(" & target.GetBranch() & ")") = False Then
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git switch " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    'ローカルブランチを削除する
    If WorkerCommon.IsExistLocalBranch(prms, target.GetBranch()) = True Then
        cmd = "git branch -D " & target.GetBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If

    'リモートブランチを削除する
    If WorkerCommon.IsExistRemoteBranch(prms, target.GetBranch()) = True Then
        cmd = "git push origin --delete " & target.GetBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If

    Common.WriteLog "DeleteBranch E"
End Sub

Private Sub DeleteTag(ByRef target As ParamTarget)
    Common.WriteLog "DeleteTag S"
    
    If FOR_TEST = False Then
        Common.WriteLog "DeleteTag E1"
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("[For TEST] タグの削除をを実行します。(" & target.GetTag() & ")") = False Then
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    If WorkerCommon.IsExistTag(prms, target.GetTag()) = True Then
        'ローカルタグを削除する
        cmd = "git tag -d " & target.GetTag()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
        'リモートタグを削除する
        cmd = "git push origin :refs/tags/" & target.GetTag()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    End If

    Common.WriteLog "DeleteTag E"
End Sub

