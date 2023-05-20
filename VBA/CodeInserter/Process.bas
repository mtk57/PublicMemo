Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String
Private Const VB6_EXT = "bas,frm,cls,ctl"
Private Const VBNET_EXT = "vb"

'�p�����[�^
Private main_param As MainParam

Private target_files() As String

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    '�Ώۃt�@�C������������
    SearchTargetFile
    
    '�Ώۃt�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
    CopyTargetFiles
    
    '���C�����[�v
    Dim i As Long
    For i = LBound(target_files) To UBound(target_files)
        Dim targer_path As String: targer_path = target_files(i)
        Common.WriteLog "i=" & i & ":[" & targer_path & "]"
    
        '�Ώۃt�@�C���̊֐��̐擪�ƍŌ�Ƀ��O�𖄂ߍ���
        InsertLog targer_path
    Next i

    Common.WriteLog "Run E"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String
    
    'Main Params
    Set main_param = New MainParam
    main_param.Init
    main_param.Validate

    Common.WriteLog main_param.GetAllValue()

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'�Ώۃt�@�C������������
Private Sub SearchTargetFile()
    Common.WriteLog "SearchTargetFile S"
    
    Dim err_msg As String
    Dim path As String
    Dim i As Long: i = 0
    
    Erase target_files
    
    '�Ώۃt�@�C������������

    '�g���q���X�g�쐬
    Dim ext_list() As String
    If main_param.GetTargetExtension() = "VB6�n" Then
        ext_list = Split(VB6_EXT, ",")
    Else
        ext_list = Split(VBNET_EXT, ",")
    End If

    '�g���q�Ń��[�v
    For i = LBound(ext_list) To UBound(ext_list)
        '�g���q�Ō������ăt�@�C�����X�g�쐬
        Dim temp_list() As String
        temp_list = Common.CreateFileList( _
                        main_param.GetTargetDirPath(), _
                        "*." & ext_list(i), _
                        main_param.IsSubDir() _
                    )
        '���ʃ}�[�W
        target_files = Common.MergeArray(target_files, temp_list)
    Next i
    
    target_files = Common.DeleteEmptyArray(target_files)
    
    If Common.IsEmptyArray(target_files) = True Then
        err_msg = "�Ώۃt�@�C����������܂���ł���"
        Err.Raise 53, , err_msg
    End If
    
    Common.WriteLog "SearchTargetFile E"
End Sub

'�Ώۃt�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
Private Sub CopyTargetFiles()
    Common.WriteLog "CopyTargetFiles S"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim base_path As String: base_path = Common.GetCommonString(target_files)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim dst_file_path() As String
    Dim i As Integer
    Dim cnt As Integer: cnt = 0
    Dim err_msg As String: err_msg = ""
    
    Common.DeleteFolder main_param.GetDestDirPath()
    
    For i = LBound(target_files) To UBound(target_files)
        Dim src As String: src = target_files(i)
        
        If Common.IsExistsFile(src) = False Then
            err_msg = "�t�@�C�������݂��܂���" & vbCrLf & _
                      "src=" & src
            Common.WriteLog "[CopyTargetFiles] �����G���[! err_msg=" & err_msg
            
            If Common.ShowYesNoMessageBox( _
                "[CopyTargetFiles]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
                "err_msg=" & err_msg _
                ) = False Then
                Err.Raise 53, , "[CopyProjectFiles] �G���[! (err_msg=" & err_msg & ")"
            End If
            
            GoTo CONTINUE
        End If
        
        'TODO: ���O�t�@�C���͏��O����
        
        'TODO: ���O�L�[���[�h���܂ރt�@�C�����͏��O����
        
        Dim dst As String: dst = main_param.GetDestDirPath() & SEP & dst_base_path & Replace(src, base_path, "")
        Dim path As String: path = Common.GetFolderNameFromPath(dst)
        
        '�t�H���_�����݂��Ȃ��ꍇ�͍쐬����
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        '�t�@�C�����R�s�[����
        fso.CopyFile src, dst
        
        ReDim Preserve dst_file_path(cnt)
        dst_file_path(cnt) = dst
        
        cnt = cnt + 1
        
CONTINUE:
        
    Next i
    
    Erase target_files
    target_files = Common.MergeArray(target_files, dst_file_path)
    target_files = Common.DeleteEmptyArray(target_files)
    
    '�N�_�t�H���_���ړ�����
    MoveBaseFolder
    
    Set fso = Nothing
    
    Common.WriteLog "CopyTargetFiles E"
End Sub

'�N�_�t�H���_���ړ�����
Private Sub MoveBaseFolder()
    Common.WriteLog "MoveBaseFolder S"

    If main_param.GetBaseDir() = "" Then
        Common.WriteLog "MoveBaseFolder E1"
        Exit Sub
    End If
    
    '�N�_�t�H���_�����w�肳��Ă���ꍇ�A�R�s�[��t�H���_�p�X�ɑ��݂��邩�`�F�b�N����
    Dim base_dir As String: base_dir = ""
    Dim i As Long
    For i = LBound(target_files) To UBound(target_files)
        base_dir = Common.GetFolderPathByKeyword( _
                        Common.GetFolderNameFromPath(target_files(i)), _
                        main_param.GetBaseDir())
        If base_dir <> "" Then
            Exit For
        End If
    Next i
    
    '���݂��Ȃ��ꍇ�͉������Ȃ�
    If base_dir = "" Then
        Common.WriteLog "MoveBaseFolder E2"
        Exit Sub
    End If
    
    Dim renamed_dir As String: renamed_dir = main_param.GetBaseDir()
    
    '���݂���ꍇ�͈ړ�����
    If Common.IsExistsFolder(main_param.GetDestDirPath() & SEP & renamed_dir) = True Then
        '�ړ���ɓ����t�H���_������ꍇ�̓��j�[�N�ȃt�H���_���ɂ���
        renamed_dir = Common.GetLastFolderName( _
                            Common.ChangeUniqueDirPath( _
                                main_param.GetDestDirPath() & SEP & renamed_dir))
    End If
    
    Common.MoveFolder base_dir, main_param.GetDestDirPath() & SEP & renamed_dir
    
    '�Ō�Ƀt�H���_���폜����
    Dim dust_dir As String: dust_dir = Replace(base_dir, main_param.GetDestDirPath() & SEP, "")
    Dim del_dir_path As String: del_dir_path = main_param.GetDestDirPath() & SEP & Split(dust_dir, SEP)(0)
    
    If Common.IsExistsFolder(del_dir_path) = False Then
        Common.WriteLog "MoveBaseFolder E3"
        Exit Sub
    End If
    
    Common.DeleteFolder del_dir_path
    
    '�Ώۃt�@�C�����X�g���č쐬����
    For i = LBound(target_files) To UBound(target_files)
        Dim new_path As String
        new_path = Replace(target_files(i), base_dir, main_param.GetDestDirPath() & SEP & renamed_dir)
        target_files(i) = new_path
    Next i
    
    Common.WriteLog "MoveBaseFolder E"
End Sub

'�Ώۃt�@�C���̊֐��̐擪�ƍŌ�Ƀ��O�𖄂ߍ���
Private Sub InsertLog(ByVal target_path As String)
    Common.WriteLog "InsertLog S"
    
    Dim contents() As String: contents = GetTargetContents(target_path)
    
    If Common.IsEmptyArray(contents) = True Then
        Common.WriteLog "InsertLog E1"
        Exit Sub
    End If
    
    Const METHOD_START = "^(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+.*\("
    Const FUNC_END = "^End Function"
    Const SUB_END = "^End Sub"
    
    Dim new_contents() As String
    Dim now_row As String
    Dim i As Long
    Dim mehod_name As String
    Dim offset As Long
    Dim is_start As Boolean: is_start = False
    Dim suffix As String
    Dim is_found_method As Boolean
    
    Dim cnt As Long: cnt = UBound(contents)
    
    For i = LBound(contents) To UBound(contents)
        now_row = contents(i)
    
        '�V�����z��Ɍ��ݍs���R�s�[����
        offset = 1
        ReDim Preserve new_contents(cnt + offset)
        new_contents(cnt + offset) = now_row
    
        '���K�\���Ŋ֐�(Function or Sub)�̎n�܂� or �I����������
        If Common.IsMatchByRegExp(now_row, METHOD_START, True) = True Then
            is_found_method = True
        ElseIf Common.IsMatchByRegExp(now_row, FUNC_END, True) = True Then
            is_start = False
        ElseIf Common.IsMatchByRegExp(now_row, SUB_END, True) = True Then
            is_start = False
        Else
            '�֐��̊J�nor�I���s�ł͂Ȃ��̂Ŏ��s��
            GoTo CONTINUE
        End If
        
        '�V�����z���1�s�ǉ����āA���O�̍s��ǉ�����
        If is_start = True Then
            suffix = " START"
        Else
            suffix = " END"
        End If
        offset = 2
        new_contents(cnt + offset) = Replace(main_param.GetInsertWord(), "��", mehod_name & suffix)
        
CONTINUE:
        cnt = cnt + offset
    Next i
    
    '�Ō�Ƀt�@�C���ɏo�͂���
    Common.CreateSJISTextFile new_contents, target_path
    
    Common.WriteLog "InsertLog E"
End Sub

Private Function GetTargetContents(ByVal path As String) As String()
    Common.WriteLog "GetTargetContents S"
    
    Dim raw_contents As String
    Dim contents() As String
    
    '�t�@�C�����J���āA�S�s��z��Ɋi�[����
    If Common.IsSJIS(path) = True Then
        raw_contents = Common.ReadTextFileBySJIS(path)
    ElseIf Common.IsUTF8(path) = True Then
        raw_contents = Common.ReadTextFileByUTF8(path)
    Else
        Dim err_msg As String: err_msg = "���T�|�[�g�̃G���R�[�h�ł�" & vbCrLf & _
                  "path=" & path
        Common.WriteLog "[GetTargetContents] �����G���[! err_msg=" & err_msg
        
        If Common.ShowYesNoMessageBox( _
            "[GetTargetContents]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetTargetContents] �G���[! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog "GetTargetContents E1"
        GetTargetContents = contents
        Exit Function
    End If
    
    contents = Split(raw_contents, vbCrLf)
    
    GetTargetContents = contents

    Common.WriteLog "GetTargetContents E"
End Function

Private Function InsertCodeForMethod( _
    ByRef contents() As String, _
    ByVal start_row As Long, _
    ByVal end_row As Long _
) As String()
    Common.WriteLog "InsertCodeForMethod S"






    Common.WriteLog "InsertCodeForMethod E"
End Function
