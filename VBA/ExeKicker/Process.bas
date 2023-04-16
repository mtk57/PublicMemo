Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Private main_param As MainParam
Private sub_params() As SubParam

'�O���[�o��
Private current_wk_src_dir_path As String
Private current_wk_dst_dir_path As String
Private before_wk_dst_dir_path As String

'���C������
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    If CheckAndCollectParam() = False Then
        Common.WriteLog "Run E1"
        Exit Sub
    End If
    
    'Sub Param�����Ɏ��s���Ă���
    If ExecSubParam() = False Then
        Common.WriteLog "Run E2"
        Exit Sub
    End If
    
    Common.WriteLog "Run E"
    MsgBox "�I���܂���"
End Sub

'��Ɨp�t�H���_�폜
Public Sub DelWkDir()
    Common.WriteLog "DelWkDir S"
    
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    If CheckAndCollectParam() = False Then
        Common.WriteLog "DelWkDir E1"
        Exit Sub
    End If
    
    DeleteWorkFolder True
    
    Common.WriteLog "DelWkDir E"
    MsgBox "�I���܂���"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Function CheckAndCollectParam() As Boolean
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    err_msg = main_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        CheckAndCollectParam = False
        Common.WriteLog "CheckAndCollectParam E1 (" & err_msg & ")"
        Exit Function
    End If
    
    Common.WriteLog main_param.GetAllValue()

    'Sub Params
    Const START_ROW = 22
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
            Common.WriteLog "CheckAndCollectParam E2 (row=" & row & ", msg=" & err_msg & ")"
            Exit Function
        End If
        
        Common.WriteLog sub_param.GetAllValue()
        
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
    Common.WriteLog "CheckAndCollectParam E"
End Function

'Sub Param�����Ɏ��s���Ă���
Private Function ExecSubParam() As Boolean
    Common.WriteLog "ExecSubParam S"
    
    Dim errmsg As String
    
    If UBound(sub_params) < 0 Then
        errmsg = "�L����Sub param������܂���"
        MsgBox errmsg
        ExecSubParam = True
        Common.WriteLog "ExecSubParam E1 (" & errmsg & ")"
        Exit Function
    End If

    Dim i, j As Integer
    Dim exe_params() As String
    Dim is_match As Boolean
    Dim is_exit_for As Boolean
    
    '�Ώۊg���q�̃t�@�C�������݂��邩�m�F����
    Dim ext As String: ext = Replace(main_param.GetInExtension(), "*", "")
    If Common.IsExistsExtensionFile(main_param.GetSrcDirPath(), ext) = False Then
        errmsg = "�����Ώۂ̊g���q�̃t�@�C�������݂��܂��� (" & ext & ")"
        MsgBox errmsg
        ExecSubParam = True
        Common.WriteLog "ExecSubParam E2 (" & errmsg & ")"
        Exit Function
    End If
    
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
            Common.WriteLog "Main Loop i=" & i & ", j=" & j
        
            '�A�Ԃ̍�Ɨp�T�u�t�H���_���쐬����
            CreateSeqWorkFolder i, j
            
            '��Ɨp�T�u�t�H���_��src��dst�ɃR�s�[����
            CopySrcToDstWorkFolder i, j
            
            'exe�ɓn���p�����[�^���X�g���쐬����
            exe_params = CreateExeParamList(sub_param)
            
            'exe�����s����
            RunExe exe_params
            
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
    
    Common.DeleteFolder (main_param.GetDestDirPath())
    Common.CopyFolder current_wk_dst_dir_path, main_param.GetDestDirPath()
    
    '��Ɨp�t�H���_���폜����
    DeleteWorkFolder main_param.IsDeleteWorkDir()

    ExecSubParam = True
    Common.WriteLog "ExecSubParam E"
End Function

'��Ɨp�t�H���_���쐬����
Private Sub CreateWorkFolder()
    Common.WriteLog "CreateWorkFolder S"

    Dim path As String: path = main_param.GetToolWorkDirPath()

    If path = "" Then
        '���w��̏ꍇ��C:\tmp�Ƃ���
        path = "C:\tmp"
        main_param.SetToolWorkDirPath (path)
    End If
    
    Common.DeleteFolder (path)

    Common.CreateFolder (path)
    
    If main_param.IsStepWorkDir() = False Then
        '�r���o�ߎc���Ȃ��ꍇ�A�Œ�T�u�t�H���_���쐬
        current_wk_src_dir_path = path & SEP & "FIX" & "_0"
        current_wk_dst_dir_path = path & SEP & "FIX" & "_1"
        Common.CreateFolder (current_wk_src_dir_path)
        Common.CreateFolder (current_wk_dst_dir_path)
    End If
    
    Common.WriteLog "CreateWorkFolder E"
End Sub

'���O���X�g�t�@�C�����쐬����
Private Sub CreateIgnoreListFile()
    Common.WriteLog "CreateIgnoreListFile S"

    If UBound(main_param.GetIgnoreFiles()) < 0 Then
        '���O�t�@�C���Ȃ�
        main_param.SetIgnoreFilePath ("")
        
        Common.WriteLog "CreateIgnoreListFile E1"
        Exit Sub
    End If
    
    '���O���X�g�t�@�C���p�X
    Const IGNORE_FILE_NAME = "ExclusionList.ini"
    Dim path As String: path = main_param.GetToolWorkDirPath() & SEP & IGNORE_FILE_NAME
    
    main_param.SetIgnoreFilePath (path)
    
    '���O���X�g�t�@�C�����쐬
    Dim filelist() As String: filelist = main_param.GetIgnoreFiles()
    Dim i As Integer
    For i = LBound(filelist) To UBound(filelist)
        Dim num As Integer: num = i + 1
        Dim ret As Integer: ret = Common.WritePrivateProfileString("SkipFile", "File" & num, filelist(i), path)
        If ret = 0 Then
            Err.Raise 53, , "���O���X�g�t�@�C���̍X�V�Ɏ��s���܂���"
        End If
    Next i
    
    Common.WriteLog "CreateIgnoreListFile E"
End Sub

'exeini���X�V
Private Sub UpdateExeIniContents(ByRef sub_param As SubParam)
    Common.WriteLog "UpdateExeIniContents S"

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
    
    Common.WriteLog "UpdateExeIniContents E"
End Sub

'�A�Ԃ̍�Ɨp�T�u�t�H���_���쐬����
Private Sub CreateSeqWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    Common.WriteLog "CreateSeqWorkFolder S"

    If main_param.IsStepWorkDir() = False Then
        Common.WriteLog "CreateSeqWorkFolder E1"
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    current_wk_src_dir_path = path & SEP & num1 & num2 & "_0"
    current_wk_dst_dir_path = path & SEP & num1 & num2 & "_1"
    Common.CreateFolder (current_wk_src_dir_path)
    Common.CreateFolder (current_wk_dst_dir_path)
    
    Common.WriteLog "CreateSeqWorkFolder E"
End Sub

'��Ɨp�T�u�t�H���_��src��dst�ɃR�s�[����
Private Sub CopySrcToDstWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    Common.WriteLog "CopySrcToDstWorkFolder S"

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
        
            Common.CopyFolder src_path, dst_path_0
            Common.CopyFolder src_path, dst_path_1
        Else
            src_path = before_wk_dst_dir_path
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
            
            Common.CopyFolder src_path, dst_path_0
            Common.CopyFolder src_path, dst_path_1
        End If
    
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
    
    If num1 = 0 And num2 = 0 Then
        '�R�s�[��͑S�t�@�C��SJIS�ɂ���!
        Common.UTF8toSJIS_AllFile current_wk_src_dir_path, main_param.GetInExtension(), main_param.IsContainSubDir()
    End If
        
    Common.WriteLog "CopySrcToDstWorkFolder E"
End Sub

'exe�ɓn���p�����[�^���X�g�쐬����
Private Function CreateExeParamList(ByRef sub_param As SubParam) As String()
    Common.WriteLog "CreateExeParamList S"

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
        
        src_path_list = Common.DeleteEmptyArray(src_path_list)
        dst_path_list = Common.DeleteEmptyArray(dst_path_list)
    End If
    
    For i = LBound(src_path_list) To UBound(src_path_list)
        ReDim Preserve param_list(i)
        '"srcdirpath" "*.vb" "dstdirpath" "inipath" "ignorefilelistpath" ""
        param_list(i) = _
            DQ & src_path_list(i) & DQ & " " & _
            DQ & main_param.GetInExtension() & DQ & " " & _
            DQ & dst_path_list(i) & DQ & " " & _
            DQ & sub_param.GetIniFilePath() & DQ & " " & _
            DQ & main_param.GetIgnoreFilePath() & DQ & " " & _
            DQ & DQ
    Next i
    
    CreateExeParamList = param_list
    
    Common.WriteLog "CreateExeParamList E"
End Function

'exe�����s����
Private Sub RunExe(ByRef param_list() As String)
    Common.WriteLog "RunExe S"

    Dim i As Integer
    Dim ret As Long
    Dim exe_param As String
    
    For i = LBound(param_list) To UBound(param_list)
        
        exe_param = _
            DQ & main_param.GetExeFilePath() & DQ & " " & _
            param_list(i)
            
        ChDir Common.GetFolderNameFromPath(main_param.GetExeFilePath())
        
        Common.WriteLog exe_param
        
        ret = Common.RunProcessWait(exe_param)
        
        If ret <> 0 Then
            Common.WriteLog "exe ret=" & ret
            Err.Raise 53, , "Exe�̎��s�Ɏ��s���܂���(ret=" & ret & ")"
        End If
    
    Next i

    Common.WriteLog "RunExe E"
End Sub

'���������邩�`�F�b�N����
Private Function IsMatch() As Boolean
    Common.WriteLog "IsMatch S"

    Dim i As Integer
    Dim is_match As Boolean: is_match = True

    '�t�@�C�����X�g���쐬
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(current_wk_src_dir_path, main_param.GetInExtension(), True)
    Dim dst_file_list() As String: dst_file_list = Common.CreateFileList(current_wk_dst_dir_path, main_param.GetInExtension(), True)

    '�t�@�C�����r
    For i = LBound(src_file_list) To UBound(src_file_list)
        is_match = Common.IsMatchTextFiles(src_file_list(i), dst_file_list(i))
        If is_match = False Then
            '1�ł����ق�����ΏI��
            IsMatch = is_match
            Common.WriteLog "IsMatch E1"
            Exit Function
        End If
    Next i
    
    IsMatch = is_match
    Common.WriteLog "IsMatch E"
End Function

'��Ɨp�T�u�t�H���_�����ւ���
Private Sub SwapWorkSubFolder()
    Common.WriteLog "SwapWorkSubFolder S"
    
    If main_param.IsStepWorkDir() = True Then
        Common.WriteLog "SwapWorkSubFolder E1"
        Exit Sub
    End If
    
    Dim tmp As String: tmp = current_wk_src_dir_path
    current_wk_src_dir_path = current_wk_dst_dir_path
    current_wk_dst_dir_path = tmp
    
    Common.WriteLog "SwapWorkSubFolder E"
End Sub

'��Ɨp�t�H���_���폜����
Private Sub DeleteWorkFolder(ByVal is_del_wk_dir As Boolean)
    Common.WriteLog "DeleteWorkFolder S"

    If is_del_wk_dir = False Then
        Common.WriteLog "DeleteWorkFolder E1"
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    
    If Common.IsExistsFolder(path) = False Then
        Common.WriteLog "DeleteWorkFolder E2"
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("��Ɨp�t�H���_���폜���܂���?") = False Then
        Common.WriteLog "DeleteWorkFolder E3"
        Exit Sub
    End If
    
    Common.DeleteFolder path
    Common.WriteLog "DeleteWorkFolder E"
End Sub
