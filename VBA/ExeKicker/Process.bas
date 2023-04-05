Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String

'�p�����[�^
Private main_param As MainParam
Private sub_params() As SubParam

'�O���[�o��
Private current_wk_src_dir_path As String
Private current_wk_dst_dir_path As String
Private before_wk_dst_dir_path As String

'���C������
Public Sub Run()
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    If CheckAndCollectParam() = False Then
        Exit Sub
    End If
    
    'Sub Param�����Ɏ��s���Ă���
    If ExecSubParam() = False Then
        Exit Sub
    End If
    
    MsgBox "�I���܂���"
End Sub

'��Ɨp�t�H���_�폜
Public Sub DelWkDir()
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    If CheckAndCollectParam() = False Then
        Exit Sub
    End If
    
    DeleteWorkFolder True
    
    MsgBox "�I���܂���"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Function CheckAndCollectParam() As Boolean
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    err_msg = main_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        CheckAndCollectParam = False
        Exit Function
    End If

    
    'Sub Params
    Const START_ROW = 21
    Const SUB_ROWS = 5
    Dim row As Integer: row = START_ROW
    Dim cnt As Integer: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        err_msg = sub_param.Init(row)
        If err_msg <> "" Then
            MsgBox err_msg
            CheckAndCollectParam = False
            Exit Function
        End If
        
        If sub_param.GetEnable() = "STOPPER" Then
            Exit Do
        ElseIf sub_param.GetEnable() = "DISABLE" Then
            GoTo CONTINUE
        End If
        
        ReDim Preserve sub_params(cnt)
        Set sub_params(cnt) = sub_param
        cnt = cnt + 1
        
CONTINUE:
        row = row + SUB_ROWS + 1
    Loop

    CheckAndCollectParam = True
End Function

'Sub Param�����Ɏ��s���Ă���
Private Function ExecSubParam() As Boolean
    If UBound(sub_params) < 0 Then
        MsgBox "�L����Sub param������܂���"
        ExecSubParam = True
        Exit Function
    End If

    Dim i, j As Integer
    Dim exe_params() As String
    Dim is_match As Boolean
    Dim is_exit_for As Boolean
    
    '��Ɨp�t�H���_���쐬����
    CreateWorkFolder
    
    '���O���X�g�t�@�C�����쐬����
    CreateIgnoreListFile
    
    For i = LBound(sub_params) To UBound(sub_params)
        is_exit_for = False
        
        Dim sub_param As SubParam
        Set sub_param = sub_params(i)
        
        'exeini���X�V����
        UpdateExeIniContents sub_param
        
        For j = 0 To main_param.GetMaxExecCount() - 1
        
            '�A�Ԃ̍�Ɨp�T�u�t�H���_���쐬����
            CreateSeqWorkFolder i, j
            
            '��Ɨp�T�u�t�H���_��src��dst�ɃR�s�[����
            CopySrcToDstWorkFolder i, j
            
            'exe�ɓn���p�����[�^���X�g���쐬����
            exe_params = CreateExeParamList(sub_param)
            
            'exe�����s����
            RunExe exe_params
            
            'TODO:���O�t�@�C���t�H�[�}�b�g�s���̂܂܂̏ꍇ�A������src��dst�ɏ��O�t�@�C�����R�s�[����
            
            '���������邩�`�F�b�N����
            is_match = IsMatch()
            If is_match = True Then
                '�S�Ĉ�v
                is_exit_for = True
            Else
                '1�ȏ�̕s��v����
                If sub_param.IsExecNotDiff() = False Then
                    is_exit_for = True
                End If
            End If
            
            before_wk_dst_dir_path = current_wk_dst_dir_path
            
            '��Ɨp�T�u�t�H���_�����ւ���
            SwapWorkSubFolder
            
            If is_exit_for = True Then
                Exit For
            End If

        Next j
    
    Next i
    
    'dst�ɃR�s�[����
    If main_param.IsStepWorkDir() = False Then
        current_wk_dst_dir_path = before_wk_dst_dir_path
    End If
    
    Common.CopyFolder current_wk_dst_dir_path, main_param.GetDestDirPath
    
    '��Ɨp�t�H���_���폜����
    DeleteWorkFolder main_param.IsDeleteWorkDir()

    ExecSubParam = True
End Function

'��Ɨp�t�H���_���쐬����
Private Sub CreateWorkFolder()
    Dim path As String: path = main_param.GetToolWorkDirPath()

    If path = "" Then
        '���w��̏ꍇ��C:\tmp�Ƃ���
        path = "C:\tmp"
        main_param.SetToolWorkDirPath (path)
    End If

    Common.CreateFolder (path)
    
    If main_param.IsStepWorkDir() = False Then
        '�r���o�ߎc���Ȃ��ꍇ�A�Œ�T�u�t�H���_���쐬
        current_wk_src_dir_path = path & SEP & "FIX" & "_0"
        current_wk_dst_dir_path = path & SEP & "FIX" & "_1"
        Common.CreateFolder (current_wk_src_dir_path)
        Common.CreateFolder (current_wk_dst_dir_path)
    End If
End Sub

'���O���X�g�t�@�C�����쐬����
Private Sub CreateIgnoreListFile()
    If UBound(main_param.GetIgnoreFiles()) < 0 Then
        '���O�t�@�C���Ȃ�
        main_param.SetIgnoreFilePath ("")
        Exit Sub
    End If
    
    '���O���X�g�t�@�C���p�X
    Const IGNORE_FILE_NAME = "TODO.ini"
    Dim path As String: path = main_param.GetToolWorkDirPath() & SEP & IGNORE_FILE_NAME
    
    main_param.SetIgnoreFilePath (path)
    
    '���O���X�g�t�@�C�����쐬
    Dim filelist() As String: filelist = main_param.GetIgnoreFiles()
    Dim i As Integer
    For i = LBound(filelist) To UBound(filelist)
        Dim num As Integer: num = i + 1
        Dim ret As Integer: ret = Common.WritePrivateProfileString("TODO", "Name" & num, filelist(i), path)
        If ret = 0 Then
            Err.Raise 53, , "���O���X�g�t�@�C���̍X�V�Ɏ��s���܂���"
        End If
    Next i
End Sub

'exeini���X�V
Private Sub UpdateExeIniContents(ByRef sub_param As SubParam)
    Dim ret As Long
    Dim path As String: path = main_param.GetExeIniFilePath()
    
    Dim addin_path As String: addin_path = main_param.GetAddinFilePath()
    ret = Common.WritePrivateProfileString("Extent", "Name1", addin_path, path)
    If ret = 0 Then
        Err.Raise 53, , "Ini�t�@�C���̍X�V�Ɏ��s���܂���(0)"
    End If

    Dim count As String
    If sub_param.IsEnableAddin = True Then
        count = "1"
    Else
        count = "0"
    End If
    ret = Common.WritePrivateProfileString("Extent", "Count", count, path)
    If ret = 0 Then
        Err.Raise 53, , "Ini�t�@�C���̍X�V�Ɏ��s���܂���(1)"
    End If
    
    Dim skip As String
    If sub_param.IsSkipComment = True Then
        skip = "1"
    Else
        skip = "0"
    End If
    ret = Common.WritePrivateProfileString("Comment", "SkipVb", skip, path)
    If ret = 0 Then
        Err.Raise 53, , "Ini�t�@�C���̍X�V�Ɏ��s���܂���(2)"
    End If
End Sub

'�A�Ԃ̍�Ɨp�T�u�t�H���_���쐬����
Private Sub CreateSeqWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    If main_param.IsStepWorkDir() = False Then
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    current_wk_src_dir_path = path & SEP & num1 & num2 & "_0"
    current_wk_dst_dir_path = path & SEP & num1 & num2 & "_1"
    Common.CreateFolder (current_wk_src_dir_path)
    Common.CreateFolder (current_wk_dst_dir_path)
End Sub

'��Ɨp�T�u�t�H���_��src��dst�ɃR�s�[����
Private Sub CopySrcToDstWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    Dim src_path As String
    Dim dst_path_0 As String
    Dim dst_path_1 As String
    
    If main_param.IsStepWorkDir() = True Then
        '�r���o�߂��c��
        
        If num1 = 0 And num2 = 0 Then
            '�ŏ������͖{����src����R�s�[����
            src_path = main_param.GetSrcDirPath()
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
        Else
            src_path = before_wk_dst_dir_path
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
        End If
        
        Common.CopyFolder src_path, dst_path_0
        Common.CopyFolder src_path, dst_path_1
    
    Else
        '�r���o�߂��c���Ȃ�
        
        If num1 = 0 And num2 = 0 Then
            '�ŏ������͖{����src����R�s�[����
            src_path = main_param.GetSrcDirPath()
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
            
            Common.CopyFolder src_path, dst_path_0
            Common.CopyFolder src_path, dst_path_1
        Else
            Common.CopyFolder current_wk_src_dir_path, current_wk_dst_dir_path
        End If
        
    End If
        
End Sub

'exe�ɓn���p�����[�^���X�g�쐬����
Private Function CreateExeParamList(ByRef sub_param As SubParam) As String()
    Dim i As Integer
    Dim param_list() As String
    
    Dim src_path_list() As String
    Dim dst_path_list() As String
    
    If main_param.IsContainSubDir() = False Then
        ReDim src_path_list(0)
        ReDim dst_path_list(0)
        src_path_list(0) = current_wk_src_dir_path
        dst_path_list(0) = current_wk_dst_dir_path
    Else
        src_path_list = Common.GetFolderPathList(current_wk_src_dir_path)
        dst_path_list = Common.GetFolderPathList(current_wk_dst_dir_path)
        
        Common.AppendArray src_path_list, current_wk_src_dir_path
        Common.AppendArray dst_path_list, current_wk_dst_dir_path
    End If
    
    For i = LBound(src_path_list) To UBound(src_path_list)
        ReDim Preserve param_list(i)
        '"srcdirpath" "dstdirpath" "*.vb" "inipath" "ignorefilelistpath" ""
        param_list(i) = _
            Chr(34) & src_path_list(i) & Chr(34) & " " & _
            Chr(34) & dst_path_list(i) & Chr(34) & " " & _
            Chr(34) & main_param.GetInExtension() & Chr(34) & " " & _
            Chr(34) & sub_param.GetIniFilePath() & Chr(34) & " " & _
            Chr(34) & main_param.GetIgnoreFilePath() & Chr(34) & " " & _
            Chr(34) & Chr(34)
    Next i
    
    CreateExeParamList = param_list
End Function

'exe�����s����
Private Sub RunExe(ByRef param_list() As String)
    Dim i As Integer
    Dim ret As Long
    Dim exe_param As String
    
    For i = LBound(param_list) To UBound(param_list)
        
        exe_param = _
            Chr(34) & main_param.GetExeFilePath() & Chr(34) & " " & _
            param_list(i)
        
        ret = Common.RunProcessWait(exe_param)
        
        If ret <> 0 Then
            Err.Raise 53, , "Exe�̎��s�Ɏ��s���܂���(ret=" & ret & ")"
        End If
    
    Next i

End Sub

'���������邩�`�F�b�N����
Private Function IsMatch() As Boolean
    Dim i As Integer
    Dim is_match As Boolean: is_match = True

    '�t�@�C�����X�g���쐬
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(current_wk_src_dir_path, main_param.GetInExtension())
    Dim dst_file_list() As String: dst_file_list = Common.CreateFileList(current_wk_dst_dir_path, main_param.GetInExtension())

    '�t�@�C�����r
    For i = LBound(src_file_list) To UBound(src_file_list)
        is_match = Common.IsMatchTextFiles(src_file_list(i), dst_file_list(i))
        If is_match = False Then
            '1�ł����ق�����ΏI��
            IsMatch = is_match
            Exit Function
        End If
    Next i
    
    IsMatch = is_match
End Function

'��Ɨp�T�u�t�H���_�����ւ���
Private Sub SwapWorkSubFolder()
    If main_param.IsStepWorkDir() = True Then
        Exit Sub
    End If
    
    Dim tmp As String: tmp = current_wk_src_dir_path
    current_wk_src_dir_path = current_wk_dst_dir_path
    current_wk_dst_dir_path = tmp
End Sub

'��Ɨp�t�H���_���폜����
Private Sub DeleteWorkFolder(ByVal is_del_wk_dir As Boolean)
    If is_del_wk_dir = False Then
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    
    If Common.IsExistsFolder(path) = False Then
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("��Ɨp�t�H���_���폜���܂���?") = False Then
        Exit Sub
    End If
    
    Common.DeleteFolder path
End Sub
