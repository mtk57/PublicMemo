Attribute VB_Name = "Process_003"
Option Explicit

Private prms As ParamContainer
Private SEP As String
Private DQ As String

Public Sub Run()
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
        
    WorkerCommon.DoClone prms
    
    For i = LBound(targetlist) To UBound(targetlist)
    
        Set target = targetlist(i)
    
        WorkerCommon.SwitchDevelopBranch prms
        
        WorkerCommon.DoPull prms
        
        CollectTag target
    
    Next i
        
    Common.WriteLog "Run E"
End Sub

Private Sub CollectTag(ByRef target As ParamTarget)
    Common.WriteLog "CollectTag S"
    
    Dim tag_list() As String
    tag_list = GetTargetTagList(target)
    
    If Common.IsEmptyArray(tag_list) = True Then
        Common.WriteLog "CollectTag E1"
        Exit Sub
    End If
    
    'タグをzipで保存する
    Dim cmd As String
    Dim git_result() As String
    Dim i As Long
    
    For i = LBound(tag_list) To UBound(tag_list)
        cmd = "git archive " & tag_list(i) & " -o " & prms.GetDestDirPath() & SEP & tag_list(i) & ".zip"
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    Next i
        
    Common.WriteLog "CollectTag E"
End Sub

Private Function GetTargetTagList(ByRef target As ParamTarget) As String()
    Common.WriteLog "GetTargetTagList S"

    Dim cmd As String
    Dim git_result() As String
    
    Dim find_word As String: find_word = target.GetTag()
    find_word = Replace(find_word, "STEP1.1", prms.GetCollectStep())
    
    'タグを検索
    cmd = "git tag --list " & find_word
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    If Common.IsEmptyArray(git_result) = True Then
        Common.WriteLog "Tag is nothing.(" & find_word & ")"
    End If
    
    GetTargetTagList = git_result

    Common.WriteLog "GetTargetTagList E"
End Function
