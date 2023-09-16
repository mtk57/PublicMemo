Attribute VB_Name = "Process_005"
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
    
    prms.SetProcessType PROCESS_TYPE.PROC_005
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
    
    '�܂��͑S�Ẵu�����`�̑��݃`�F�b�N
    For i = 0 To UBound(targetlist)
    
        Set target = targetlist(i)
    
        If WorkerCommon.IsExistBranch(prms, target.GetBranch()) = False Then
            msg = "�u�����`��������܂���B(" & target.GetBranch() & ")"
        
        ElseIf WorkerCommon.IsExistTag(prms, target.GetTag(), True) = True Or _
               WorkerCommon.IsExistTag(prms, target.GetTag(), False) = True Then
            msg = "�^�O�����łɃ��[�J���܂��̓����[�g�ɑ��݂��Ă��܂��B(tag=" & target.GetTag() & ")"
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
        Common.WriteLog "targetlist_exist_only is empty."
        GoTo FINISH
    End If
    
    For i = 0 To UBound(targetlist_exist_only)
    
        Set target = targetlist_exist_only(i)
        
        WorkerCommon.SwitchDevelopBranch prms
        
        WorkerCommon.DoPull prms
    
        WorkerCommon.SwitchBranch prms, target
        
        WorkerCommon.DoPull prms
        
        WorkerCommon.DoMerge prms, prms.GetBaseBranch()
        
        RunSakura prms, target
        
        If DoCommit(target) = False Then
            '�㑱�������X�L�b�v
            GoTo CONTINUE2
        End If
        
        DoTag target

        DoPush target.GetBranch()
        
        WorkerCommon.SwitchDevelopBranch prms
    
        WorkerCommon.DoMerge prms, target.GetBranch()
        
CONTINUE2:
        
    Next i
    
    DoPush prms.GetBaseBranch()
        
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

Private Sub DoPush(ByVal branch As String)
    Common.WriteLog "DoPush S"
    
    If prms.IsUpdateRemote() = False Then
        Common.WriteLog "DoPush E1"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    '�^�O��t����
    cmd = "git push -f --tags --set-upstream origin " & branch
    
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

Private Function DoCommit(ByRef target As ParamTarget) As Boolean
    Common.WriteLog "DoCommit S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�R�~�b�g����
    cmd = "git commit -a -m " & DQ & target.GetCommit() & DQ
    
On Error Resume Next
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Dim err_msg As String: err_msg = Err.Description
    Err.Clear
On Error GoTo 0

    If err_msg = "" Then
        '�����B�㑱�������s�t���O��ON
        DoCommit = True
    ElseIf InStr(err_msg, "exit code=1") = 0 Then
        'exit code=1�ȊO�͏�ʂɍēx�G���[�ʒm
        Err.Raise 53, , "[DoCommit] git commit�ŃG���[ (err_msg=" & err_msg & ")"
    Else
        'exit code=1�ł����s�ł���\����������̂͑��s����
        
        If InStr(err_msg, "working tree clean") = 0 Then
            '���s�ł��Ȃ����̂Ƃ���B��ʂɍēx�G���[�ʒm
            Err.Raise 53, , "[DoCommit] git commit�ŃG���[ (err_msg=" & err_msg & ")"
        End If
        
        DoCommit = False  '�R�~�b�g�������̂��Ȃ��̂Ō㑱�������s�t���O��OFF
    End If
    
    Common.WriteLog "DoCommit E"
End Function

Private Sub RunSakura(ByRef prms As ParamContainer, ByVal target As ParamTarget)
    Common.WriteLog "RunSakura S"
    
    Dim ref_file_list() As String
    
    '�Ώ�VB�v���W�F�N�g�̃t�@�C���݂̂�sakura�ŏ������邽�߁A��Ɨp�t�H���_�[�ɃR�s�[
    Dim wk_dir As String: wk_dir = CopyVBProjectFilesToWorkDir(target, ref_file_list)
    
    Dim ret As Long
    Dim sakura_param As String
    sakura_param = prms.GetSakuraPath() & " " & CreateSakuraArgs(prms, wk_dir)

    Common.WriteLog "sakura_param=" & sakura_param
    
    ret = Common.RunBatFile(sakura_param)
  
    If ret <> 0 Then
        Common.DeleteFolder wk_dir
        Err.Raise 53, , "[RunSakura] sakura�̎��s�Ɏ��s���܂����B(sakura_param=" & sakura_param & ", ret=" & ret & ")"
    End If
    
    Dim base_dir As String: base_dir = Common.GetCommonString(ref_file_list)
    
    CopyVBProjectFilesFromWorkDir wk_dir, base_dir
    Common.DeleteFolder wk_dir
  
    Common.WriteLog "RunSakura E"
End Sub

Private Function CopyVBProjectFilesToWorkDir(ByVal target As ParamTarget, ByRef ref_file_list() As String) As String
    Common.WriteLog "CopyVBProjectFilesToWorkDir S"
    
    Dim SEP As String: SEP = Application.PathSeparator
    Dim err_msg As String
    Dim vbproj_path As String: vbproj_path = target.GetVBPrjFilePath()
    
    If Common.IsExistsFile(vbproj_path) = False Then
        err_msg = "[CopyVBProjectFilesToWorkDir] File not found.(" & vbproj_path & ")"
        If Common.ShowYesNoMessageBox( _
            "VB�v���W�F�N�g�t�@�C�������݂��܂���B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "Error! (err_msg=" & err_msg & ")"
        End If
    End If
    
    'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
    Dim raw_contents As String: raw_contents = Common.ReadTextFileBySJIS(vbproj_path)
    
    '�t�@�C���̓��e��z��Ɋi�[����
    Dim contents() As String: contents = Split(raw_contents, vbCrLf)
    contents = Common.DeleteEmptyArray(contents)

    If Common.IsEmptyArray(contents) = True Then
        err_msg = "[CopyVBProjectFilesToWorkDir] VB Project file is empty.(" & vbproj_path & ")"
        If Common.ShowYesNoMessageBox( _
            "VB�v���W�F�N�g�t�@�C������ł��B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "Error! (err_msg=" & err_msg & ")"
        End If
    End If
    
    Dim ext As String: ext = Common.GetFileExtension(vbproj_path)
    
    'VB�v���W�F�N�g�t�@�C������͂��āA�Q�Ƃ���Ă���t�@�C���̈ꗗ���擾����
    If ext = "vbp" Then
        'VB6
        ref_file_list = WorkerCommon.GetRefFileListForVB6project(prms, vbproj_path, contents)
    Else
        'VB.NET
        ref_file_list = WorkerCommon.GetRefFileListForVBdotNetProject(prms, vbproj_path, contents)
    End If
    
    '��Ɨp�t�H���_���쐬���ăR�s�[
    Dim wk_dir As String: wk_dir = GetTempFolder() & SEP & GetNowTimeString()
    
    Dim base_dir As String: base_dir = Common.GetCommonString(ref_file_list)
    Dim base_parent As String: base_parent = Common.GetLastFolderName(base_dir)
    Dim i As Long
    For i = 0 To UBound(ref_file_list)
        Dim dst_dir As String: dst_dir = Common.GetFolderNameFromPath(wk_dir & SEP & base_parent & SEP & Replace(ref_file_list(i), base_dir, ""))
        Common.CreateFolder dst_dir
        Common.CopyFile ref_file_list(i), dst_dir & SEP & Common.GetFileName(ref_file_list(i))
    Next i
    
    CopyVBProjectFilesToWorkDir = wk_dir
    Common.WriteLog "CopyVBProjectFilesToWorkDir E"
End Function

Private Sub CopyVBProjectFilesFromWorkDir(ByVal wk_dir As String, ByVal base_dir As String)
    Common.WriteLog "CopyVBProjectFilesFromWorkDir S"
    
    Common.CopyFolder wk_dir & Application.PathSeparator & Common.GetLastFolderName(base_dir), base_dir
    
    Common.WriteLog "CopyVBProjectFilesFromWorkDir E"
End Sub

Private Function CreateSakuraArgs(ByRef prms As ParamContainer, ByVal wk_dir As String) As String
    Common.WriteLog "CreateSakuraArgs S"
    
    CreateSakuraArgs = ""
    
    If prms.GetSakuraArgs() = "" Then
        Common.WriteLog "CreateSakuraArgs E1"
        Exit Function
    End If
    
    Dim args() As String
    args = Split(prms.GetSakuraArgs(), vbLf)
    
    Dim i As Long
    For i = 0 To UBound(args)
        If InStr(args(i), "-GFOLDER=") > 0 Then
            args(i) = "-GFOLDER=" & DQ & wk_dir & DQ
            Exit For
        End If
    Next i
    
    CreateSakuraArgs = Join(args, " ")
    
    Common.WriteLog "CreateSakuraArgs E"
End Function
