Attribute VB_Name = "Process_004"
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
    
    prms.SetProcessType PROCESS_TYPE.PROC_004
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    Dim i As Long
    Dim cnt As Long: cnt = 0
    Dim target As ParamTarget
    Dim targetlist() As ParamTarget
    Dim targetlist_exist_only() As ParamTarget
    targetlist = prms.GetTargetList()
        
    WorkerCommon.DoClone prms
    
    '�܂��͑S�Ẵu�����`�̑��݃`�F�b�N�ƃ^�O�`�F�b�N
    For i = 0 To UBound(targetlist)
    
        Set target = targetlist(i)
    
        If WorkerCommon.IsExistBranch(prms, target.GetBranch()) = False Then
            msg = "�u�����`��������܂���B(" & target.GetBranch() & ")"
        ElseIf InStr(target.GetTag(), "STEP1.8") = 0 Then
            msg = "�^�O��STEP1.8���w�肳��Ă��܂���B (tag=" & target.GetTag() & ")"
        End If

        If msg <> "" Then
            Common.WriteLog msg
            If Common.ShowYesNoMessageBox( _
                "�����O�`�F�b�N�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
                "err_msg=" & msg _
                ) = False Then
                Err.Raise 53, , "[Run] �G���[ (err_msg=" & msg & ")"
            End If
            GoTo CONTINUE
        End If
        
        msg = ""

        ReDim Preserve targetlist_exist_only(cnt)
        Set targetlist_exist_only(cnt) = target
        cnt = cnt + 1
            
CONTINUE:
            
    Next i
    
    If Common.IsEmptyArray(targetlist_exist_only) = True Then
        GoTo FINISH
    End If
    
    For i = 0 To UBound(targetlist_exist_only)
    
        Set target = targetlist_exist_only(i)
    
        WorkerCommon.SwitchBranch prms, target
        
        WorkerCommon.DoPull prms
        
        DoTag target
        
        DoPush target
    
    Next i
        
FINISH:
        
    Common.WriteLog "Run E"
End Sub

Private Sub DoTag(ByRef target As ParamTarget)
    Common.WriteLog "DoTag S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�^�O��t����
    cmd = "git tag -f " & target.GetTag() & " HEAD"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoTag E"
End Sub

Private Sub DoPush(ByRef target As ParamTarget)
    Common.WriteLog "DoPush S"
    
    If prms.IsUpdateRemote() = False Then
        Common.WriteLog "DoPush E1"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    '�^�O��t����
    cmd = "git push -f --tags --set-upstream origin " & target.GetBranch()
    
On Error Resume Next
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Dim err_msg As String: err_msg = Err.Description
    Err.Clear
On Error GoTo 0

    If err_msg = "" Then
        '����
    ElseIf InStr(err_msg, "exit code=1") = 0 Then
        'exit code=1�ȊO�͏�ʂɍēx�G���[�ʒm
        Err.Raise 53, , "[DoPush] git push�ŃG���[ (err_msg=" & err_msg & ")"
    Else
        'exit code=1�͑��s�ł���\���������̂Ŋm�F
        If Common.ShowYesNoMessageBox( _
            "git push�ňȉ��̃G���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[DoPush] git push�ŃG���[ (err_msg=" & err_msg & ")"
        End If
        Common.WriteLog err_msg
    End If
    
    Common.WriteLog "DoPush E"
End Sub





