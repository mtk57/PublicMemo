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
    
    prms.SetProcessType PROCESS_TYPE.PROC_003
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    If Common.IsExistsFolder(prms.GetDestDirPath) = False Then
        Common.CreateFolder prms.GetDestDirPath()
    End If
    
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

'�^�O�����W����
Private Sub CollectTag(ByRef target As ParamTarget)
    Common.WriteLog "CollectTag S"
    
    Dim tag_list() As String
    tag_list = GetTargetTagList(target)
    
    If Common.IsEmptyArray(tag_list) = True Then
        Common.WriteLog "Tag not found.(Branch=" & target.GetBranch() & ", Tag=" & target.GetTag() & ")"
        Common.WriteLog "CollectTag E1"
        Exit Sub
    End If
    
    '�^�O��zip�ŕۑ�����
    Dim i As Long
    For i = LBound(tag_list) To UBound(tag_list)
        DoArchive target, tag_list(i)
    Next i
        
    Common.WriteLog "CollectTag E"
End Sub

Private Function GetTargetTagList(ByRef target As ParamTarget) As String()
    Common.WriteLog "GetTargetTagList S"

    Dim cmd As String
    Dim git_result() As String
    
    Dim find_word As String: find_word = target.GetTag()
    
    '�^�O������
    cmd = "git tag --list " & find_word
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    If Common.IsEmptyArray(git_result) = True Then
        Common.WriteLog "Tag not found.(" & find_word & ")"
    End If
    
    GetTargetTagList = git_result

    Common.WriteLog "GetTargetTagList E"
End Function

'�^�O��zip�ŕۑ�����
Private Sub DoArchive(ByVal target As ParamTarget, ByVal tag As String)
    Common.WriteLog "DoArchive S"
    Common.WriteLog "��tag=" & tag

    If prms.IsIgnoreNotRef() = False Then
        GoTo FINISH
    End If

    Dim i As Long  '0=vbp, 1=vbproj
    
    Dim target_filename As String: target_filename = Common.GetFileName(target.GetVBPrjFilePath())
    Dim target_filename_list(1) As String
    
    If Common.GetFileExtension(target_filename) = "vbp" Then
        target_filename_list(0) = target_filename
        target_filename_list(1) = Replace(target_filename_list(0), ".vbp", ".vbproj")
    Else
        target_filename_list(1) = target_filename
        target_filename_list(0) = Replace(target_filename_list(1), ".vbproj", ".vbp")
    End If
    
    Dim ref_file_list() As String
    Dim ref_file_list_vb6() As String
    Dim ref_file_list_vbnet() As String
    
    For i = LBound(target_filename_list) To UBound(target_filename_list)
        Dim filepath As String: filepath = WorkerCommon.GetFilepathByTag(prms, tag, target_filename_list(i))
        
        If filepath = "" Then
            GoTo CONTINUE
        End If
        
        Dim contents() As String: contents = WorkerCommon.DoShow(prms, tag & ":" & filepath)
        
        'VB�v���W�F�N�g�t�@�C������͂��āA�Q�Ƃ���Ă���t�@�C���̈ꗗ���擾����
        If i = 0 Then
            'VB6
            ref_file_list_vb6 = GetRefFileListForVB6project(filepath, contents)
        Else
            'VB.NET
            ref_file_list_vbnet = GetRefFileListForVBdotNetProject(filepath, contents)
        End If
    
CONTINUE:
    
    Next i
    
    ref_file_list = Common.DeleteEmptyArray(Common.VariantToStringArray(Common.MergeArray(ref_file_list_vb6, ref_file_list_vbnet)))
    
    If Common.IsEmptyArray(ref_file_list) = True Then
        Common.WriteLog "ref_file_list is empty."
        Common.WriteLog "DoArchive E1"
        Exit Sub
    End If
    
FINISH:
    Dim cmd As String
    Dim git_result() As String
    Dim zip_file As String: zip_file = prms.GetDestDirPath() & SEP & tag & ".zip"
    
    If prms.IsIgnoreNotRef() = False Then
        cmd = "git archive " & tag & " -o " & zip_file
    Else
        'git archive --format=zip --output=<�o�̓t�@�C����>.zip <�^�O��> <�t�@�C���p�X>
        Dim files As String: files = Common.JoinFromArray(ref_file_list, " ", True)
        cmd = "git archive --format=zip --output=" & zip_file & " " & tag & " " & files
    End If
    
On Error Resume Next
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Dim err_msg As String: err_msg = Err.Description
    Err.Clear
On Error GoTo 0

    If err_msg = "" Then
        '����
    Else
        Common.WriteLog "[DoArchive] �����G���[! err_msg=" & err_msg
        If Common.ShowYesNoMessageBox( _
            "git archive�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[DoArchive] git archive�ŃG���[ (err_msg=" & err_msg & ")"
        End If
    End If

    Common.WriteLog "DoArchive E"
End Sub

'VB�v���W�F�N�g�t�@�C������͂��āA�Q�Ƃ���Ă���t�@�C���̈ꗗ���擾����
Private Function GetRefFileListForVB6project( _
    ByVal vbprj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "GetRefFileListForVB6project S"
    
    Dim ref_files() As String

    'vbp�t�@�C���ɋL�ڂ���Ă���t�@�C�������X�g�ɒǉ�
    ref_files = WorkerCommon.ParseVB6Project( _
                    prms, _
                    prms.GetGitDirPath() & SEP & Replace(vbprj_path, "/", "\"), _
                    contents _
                )
    
    '���΃p�X�ɕύX����
    ref_files = UpdateRefFiles(ref_files)

    'vbp�t�@�C�������X�g�ɒǉ�����
    Dim cnt As Long: cnt = UBound(ref_files)
    ReDim Preserve ref_files(cnt + 1)
    ref_files(cnt + 1) = vbprj_path

    GetRefFileListForVB6project = ref_files
    
    Common.WriteLog "GetRefFileListForVB6project E"
End Function

'VB.NET�v���W�F�N�g�t�@�C������͂��āA�Q�Ƃ���Ă���t�@�C���̈ꗗ���擾����
Private Function GetRefFileListForVBdotNetProject( _
    ByVal vbprj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "GetRefFileListForVBdotNetProject S"
    
    Dim ref_files() As String

    'vbproj�t�@�C���ɋL�ڂ���Ă���t�@�C�������X�g�ɒǉ�
    ref_files = WorkerCommon.ParseVBNETProject( _
                    prms, _
                    prms.GetGitDirPath() & SEP & Replace(vbprj_path, "/", "\"), _
                    contents _
                )
    
    '���΃p�X�ɕύX����
    ref_files = UpdateRefFiles(ref_files)
    
    'vbproj�t�@�C����sln�t�@�C�������X�g�ɒǉ�����
    Dim cnt As Long: cnt = UBound(ref_files)
    ReDim Preserve ref_files(cnt + 2)
    ref_files(cnt + 1) = vbprj_path
    ref_files(cnt + 2) = Replace(vbprj_path, ".vbproj", ".sln")
        
    GetRefFileListForVBdotNetProject = ref_files
    
    Common.WriteLog "GetRefFileListForVBdotNetProject E"
End Function

'���΃p�X�ɕύX����
Private Function UpdateRefFiles(ByRef ref_files() As String) As String()
    Common.WriteLog "UpdateRefFiles S"
    
    Dim ret_files() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    
    For i = LBound(ref_files) To UBound(ref_files)
        ReDim Preserve ret_files(cnt)
        ret_files(cnt) = Replace(Replace(ref_files(i), prms.GetGitDirPath() & SEP, ""), SEP, "/")
        cnt = cnt + 1
    Next i
    
    UpdateRefFiles = ret_files

    Common.WriteLog "UpdateRefFiles E"
End Function

