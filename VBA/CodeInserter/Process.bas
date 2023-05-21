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
    
        '�Ώۃt�@�C���̊֐��̐擪�ƍŌ�ɃR�[�h�𖄂ߍ���
        InsertCode targer_path
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
        
        If IsIgnoreFile(src) = True Then
            '���O�t�@�C���͏��O����
            Common.WriteLog "���O=" & src
            GoTo CONTINUE
        End If
        
        If IsIgnoreKeyword(src) = True Then
            '���O�L�[���[�h���܂ނ̂ŏ��O����
            Common.WriteLog "���O=" & src
            GoTo CONTINUE
        End If
        
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

'���O�t�@�C������Ԃ�
Private Function IsIgnoreFile(ByVal path As String) As Boolean
    Common.WriteLog "IsIgnoreFile S"
    
    If main_param.GetIgnoreFiles() = "" Then
        IsIgnoreFile = False
        Common.WriteLog "IsIgnoreFile E1"
        Exit Function
    End If

    '���O�t�@�C�����X�g���쐬
    Dim ignore_files() As String
    ignore_files = Split(main_param.GetIgnoreFiles(), ",")

    If Common.IsEmptyArray(ignore_files) = True Then
        IsIgnoreFile = False
        Common.WriteLog "IsIgnoreFile E2"
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_files) To UBound(ignore_files)
        If Common.GetFileName(path) = ignore_files(i) Then
            IsIgnoreFile = True
            Common.WriteLog "IsIgnoreFile E3"
            Exit Function
        End If
    Next i
    
    IsIgnoreFile = False
    Common.WriteLog "IsIgnoreFile E"
End Function

'���O�L�[���[�h���܂ނ���Ԃ�
Private Function IsIgnoreKeyword(ByVal path As String) As Boolean
    Common.WriteLog "IsIgnoreKeyword S"
    
    If main_param.GetIgnoreKeywords() = "" Then
        IsIgnoreKeyword = False
        Common.WriteLog "IsIgnoreKeyword E1"
        Exit Function
    End If

    '���O�t�@�C�����X�g���쐬
    Dim ignore_keywords() As String
    ignore_keywords = Split(main_param.GetIgnoreKeywords(), ",")

    If Common.IsEmptyArray(ignore_keywords) = True Then
        IsIgnoreKeyword = False
        Common.WriteLog "IsIgnoreKeyword E2"
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_keywords) To UBound(ignore_keywords)
        If InStr(Common.GetFileName(path), ignore_keywords(i)) > 0 Then
            IsIgnoreKeyword = True
            Common.WriteLog "IsIgnoreKeyword E3"
            Exit Function
        End If
    Next i
    
    IsIgnoreKeyword = False
    Common.WriteLog "IsIgnoreKeyword E"
End Function

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

'�Ώۃt�@�C���̊֐��̐擪�ƍŌ�ɃR�[�h�𖄂ߍ���
Private Sub InsertCode(ByVal target_path As String)
    Common.WriteLog "InsertCode S"
    
    Dim contents() As String: contents = GetTargetContents(target_path)
    
    If Common.IsEmptyArray(contents) = True Then
        Common.WriteLog "InsertCode E1"
        Exit Sub
    End If
    
    Const METHOD_START = "(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+.*\("

    Dim new_contents() As String
    ReDim new_contents(0)
    Dim i As Long
 
    For i = LBound(contents) To UBound(contents)
        Dim line As String: line = contents(i)
        
        If Common.IsCommentCode(line, Common.GetFileExtension(target_path)) = True Then
            '�R�����g�s�Ȃ̂Ŏ��̍s��
            GoTo NOT_METHOD
        End If
        
        If Common.IsMatchByRegExp(line, METHOD_START, True) = True Then
            '�֐���`�̊J�n�s�𔭌�
            i = i + InsertCodeForMethod( _
                        target_path, _
                        i, _
                        contents, _
                        new_contents _
                    )
            
            GoTo CONTINUE
        End If
            
NOT_METHOD:
        '�֐���`�ȊO�̍s
        Common.AppendArray new_contents, line
        
CONTINUE:
    
    Next i
    
    '�Ō�Ƀt�@�C���ɏo�͂���
    Common.CreateSJISTextFile new_contents, target_path
    
FINISH:
    Common.WriteLog "InsertCode E"

End Sub

'�֐��ɃR�[�h��}������
Private Function InsertCodeForMethod( _
    ByVal target_path As String, _
    ByVal start As Long, _
    ByRef contents() As String, _
    ByRef new_contents() As String _
) As Long
    Common.WriteLog "InsertCodeForMethod S"
    
    Const METHOD_END = "End\s(Function|Sub)"
    
    Dim i As Long
    Dim line As String: line = contents(start)  '��͒��̍s�f�[�^
    Dim start_clm As Long   '�֐���`(�J�n�s)�̌�(1�n�܂�)
    Dim end_clm As Long     '�֐���`(�I���s)�̌�(1�n�܂�)
    Dim method_name As String: method_name = GetMethodName(line)
    Dim cnt As Long     '��͂�i�߂��s���B�������J�n�s����ђǉ��s�͊܂܂Ȃ��B
    Dim offset As Long  '�֐��J�n�ʒu�̃I�t�Z�b�g�s��(�֐��̈����������s�̏ꍇ��2�s�ȏ�ɂȂ�)

    '���ʒu��ێ�
    start_clm = Common.FindFirstCasePosition(line)

    Common.AppendArray new_contents, line
    
    '�֐��J�n��`�̏I���s���擾����
    offset = GetMethodStartOffset(target_path, start, contents)
    
    If offset > 0 Then
        For i = 0 To offset
            Common.AppendArray new_contents, contents(start + 1 + i)
        Next i
    End If
    Common.AppendArray new_contents, GetMethodStartLine(method_name)
    
    For i = start + offset To UBound(contents)
        line = contents(i)
        
        If Common.IsCommentCode(line, Common.GetFileExtension(target_path)) = True Then
            '�R�����g�s�Ȃ̂Ŏ��̍s��
            GoTo METHOD_BODY
        End If
        
        If Common.IsMatchByRegExp(line, METHOD_END, True) = True Then
            '�֐���`�̏I���s�𔭌�
            
            '���ʒu��ێ�
            end_clm = Common.FindFirstCasePosition(line)
            
            If start_clm <> end_clm Then
                '�J�n���ƈقȂ�̂Ŏ��̍s��
                GoTo METHOD_BODY
            End If
            
            Common.AppendArray new_contents, GetMethodEndLine(method_name)
            Common.AppendArray new_contents, line
            cnt = cnt + 1
            
            GoTo FINISH
        End If

METHOD_BODY:
        '�֐���`�̖{��
        Common.AppendArray new_contents, line
        cnt = cnt + 1
        
    Next i

FINISH:
    InsertCodeForMethod = cnt
    Common.WriteLog "InsertCodeForMethod E"
End Function

'�֐��J�n��`�̏I���s���擾����
'
' <�l����>
'  1." "��Split
'  2."Sub"�������Sub���[�hON, "Function"�������Function���[�hON   �����̃��\�b�h�ɓn���O�ɐ��K�\���Ńq�b�g�����������n���Ă���̂Ŗ������Ƃ͗L�蓾�Ȃ��B
'  3.�s���[�v�J�n
'  4.  �񃋁[�v�J�n
'  4-1.  ��I�[��"_"������΍s���[�v���s�B�Ȃ���Ώ����I���B�s���[�v�����񐔂��I���s�Ƃ���B
'  4-2.  "("������Ί��ʃJ�E���^++�A")"������Ί��ʃJ�E���^--
'  4-3.  ���ʃJ�E���^��0�ɂȂ��� && Sub���[�h�Ȃ�Sub�̏I���Ɣ��f�������I���B�s���[�v�����񐔂��I���s�Ƃ���B
'  4-4.  ���ʃJ�E���^��0�ɂȂ��� && Function���[�h�Ȃ�Function�̈����̏I���Ɣ��f���邪�߂�l������\��������̂Ŗ߂�l���[�hON���ď������s�B
'  4-5.  �߂�l���[�h && " As "������Ζ߂�l������B��I�[��"_"���Ȃ���Ώ����I���B�s���[�v�����񐔂��I���s�Ƃ���B
'
'  - �����s�̏ꍇ�A"_"�ȍ~�̗�ɃR�����g��" "�͕t�����Ȃ��B
'  - �����s�̏ꍇ�A�R�����g�͈�ؕt�����Ȃ��B
'  - �������߂�l������Function���쐬�\
'  - �߂�l���z��̏ꍇ"()"�ŏI���̂ŁAFunction���[�h���̊��ʃJ�E���^�ɂ͒��ӂ��K�v�B
'  - ��L��<�l����>�͐���P�[�X�̂�(�܂萳��Ƀr���h�ł���R�[�h)�B
'
Private Function GetMethodStartOffset( _
    ByVal target_path As String, _
    ByVal start As Long, _
    ByRef contents() As String _
) As Long
    Common.WriteLog "GetMethodStartOffset S"
    
    Dim r As Long
    Dim c As Long
    Dim line As String
    
    
    For r = start To UBound(contents)
        line = contents(i)
        
        If Common.IsCommentCode(line, Common.GetFileExtension(target_path)) = True Or _
           Right(line, 1) = "_" Then
            '�R�����g�s or �����s�Ȃ̂Ŏ��̍s��
            GoTo CONTINUE
        End If
        
        Dim tmp As String: tmp = Common.RemoveRightComment(line, Common.GetFileExtension(target_path))
        
        If Right(RTrim(tmp), 1) = ")" Or InStr(tmp, "As") > 0 Then
            '�֐��J�n��`�̏I���s�𔭌�
            GetMethodStartOffset = i
            Common.WriteLog "GetMethodStartOffset E1"
            Exit Function
        End If

CONTINUE:
        
    Next r

    '�֐��J�n��`�̏I���s��������Ȃ�!
    GetMethodStartOffset = 0
    Common.WriteLog "GetMethodStartOffset E"
End Function

'�֐��J�n����ɑ}������R�[�h���쐬����
Private Function GetMethodStartLine(ByVal method_name As String) As String
    Common.WriteLog "GetMethodStartLine S"
    GetMethodStartLine = Replace(main_param.GetInsertWord(), "��", method_name & " START")
    Common.WriteLog "GetMethodStartLine E"
End Function

'�֐��I�����O�ɑ}������R�[�h���쐬����
Private Function GetMethodEndLine(ByVal method_name As String) As String
    Common.WriteLog "GetMethodEndLine S"
    GetMethodEndLine = Replace(main_param.GetInsertWord(), "��", method_name & " END")
    Common.WriteLog "GetMethodEndLine E"
End Function

'�֐�����Ԃ�
Private Function GetMethodName(ByVal line As String) As String
    Common.WriteLog "GetMethodName S"
    
    Const METHOD = "\s*(Function|Sub)\s+.*\("
    
    Dim list() As String
    list = Common.GetMatchByRegExp(line, METHOD, True)
    
    list = Common.DeleteEmptyArray(list)
    list = Split(list(0), " ")
    
    Dim last As Integer: last = UBound(list)
    GetMethodName = Replace(list(last), "(", "")
    
    Common.WriteLog "GetMethodName E"
End Function

'�Ώۃt�@�C����ǂݍ���œ��e��z��ŕԂ�
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


