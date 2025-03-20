Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String
'Private Const VB6_EXT = "bas,frm,cls,ctl"
'Private Const VBNET_EXT = "vb"
Private Const PLSQL_EXT = "sql"
'Private Const REPLACE_METHOD = "��"
'Private Const REPLACE_FILENAME = "��"

'Private Const METHOD = "\s*(Function|Sub)\s+.*"
'Private Const METHOD_START = "(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+.*\("
'Private Const METHOD_END = "End\s(Function|Sub)$"
'Private Const METHOD_EXIT = "Exit\s(Function|Sub)$"
'Private Const METHOD_APP_END = "^(\t|\s)*\bEnd$"
'Private Const METHOD_RET = "^[ \t]*(Return|Throw) *"

'Private Const IGNORE_WORDS = "Declare,PtrSafe,Lib,Alias"

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
    
        '�Ώۃt�@�C���̃R�����g���폜����
        DeleteComment targer_path
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
    If main_param.GetTargetExtension() = "PL/SQL(sql)" Then
        ext_list = Split(PLSQL_EXT, ",")
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
        
        If main_param.IsIgnoreFile(src) = True Then
            '���O�t�@�C���͏��O����
            Common.WriteLog "���O�t�@�C��=" & src
            GoTo CONTINUE
        End If
        
        If main_param.IsIgnoreKeyword(src) = True Then
            '���O�L�[���[�h���܂ނ̂ŏ��O����
            Common.WriteLog "���O�t�@�C��=" & src
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

'�Ώۃt�@�C���̃R�����g���폜����
Private Sub DeleteComment(ByVal target_path As String)
    Common.WriteLog "InsertCode S"
    
    Dim contents() As String: contents = GetTargetContents(target_path)
    
    If Common.IsEmptyArrayLong(contents) = True Then
        Common.WriteLog "DeleteComment E1"
        Exit Sub
    End If
    
    Dim new_contents() As String
    
    If main_param.GetTargetExtension() = "PL/SQL(sql)" Then
        'PL/SQL�R�����g�̍폜
        new_contents = RemovePLSQLComments(contents)
    End If
    
    '�Ō�Ƀt�@�C���ɏo�͂���
    Common.CreateSJISTextFile new_contents, target_path
    
FINISH:
    Common.WriteLog "DeleteComment E"

End Sub

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

Private Function RemovePLSQLComments(sqlLines() As String) As String()
    Common.WriteLog "RemovePLSQLComments S"

    Dim result() As String
    Dim resultCount As Long
    Dim inString As Boolean
    Dim inBlockComment As Boolean
    Dim i As Long, j As Long
    Dim currentChar As String
    Dim nextChar As String
    Dim currentLine As String
    Dim processedLine As String
    
    ' ���ʔz��̏�����
    ReDim result(LBound(sqlLines) To UBound(sqlLines))
    resultCount = 0
    
    inString = False
    inBlockComment = False
    
    ' �e�s������
    For i = LBound(sqlLines) To UBound(sqlLines)
        currentLine = sqlLines(i)
        processedLine = ""
        
        j = 1
        Do While j <= Len(currentLine)
            currentChar = Mid(currentLine, j, 1)
            
            If j < Len(currentLine) Then
                nextChar = Mid(currentLine, j + 1, 1)
            Else
                nextChar = ""
            End If
            
            ' �����񃊃e�����̏���
            If currentChar = "'" And Not inBlockComment Then
                ' �G�X�P�[�v���ꂽ�N�H�[�g('')�̏���
                If nextChar = "'" Then
                    processedLine = processedLine & "''"
                    j = j + 2
                    GoTo ContinueLoop
                Else
                    inString = Not inString
                    processedLine = processedLine & currentChar
                End If
            
            ' �u���b�N�R�����g�J�n�̏���
            ElseIf currentChar = "/" And nextChar = "*" And Not inString And Not inBlockComment Then
                inBlockComment = True
                j = j + 1 ' ���̕���(*)���X�L�b�v
            
            ' �u���b�N�R�����g�I���̏���
            ElseIf currentChar = "*" And nextChar = "/" And inBlockComment Then
                inBlockComment = False
                j = j + 1 ' ���̕���(/)���X�L�b�v
            
            ' �s�R�����g�̏���
            ElseIf currentChar = "-" And nextChar = "-" And Not inString And Not inBlockComment Then
                ' �s�R�����g�̏ꍇ�A���̍s�̎c��͖���
                Exit Do
            
            ' �R�����g���łȂ���Ε��������ʂɒǉ�
            ElseIf Not inBlockComment Then
                processedLine = processedLine & currentChar
            End If
            
            j = j + 1
ContinueLoop:
        Loop
        
        ' �������ꂽ�s�����ʔz��ɒǉ�
        result(resultCount) = processedLine
        resultCount = resultCount + 1
    Next i
    
    ' ���ʔz��̃T�C�Y�����ۂ̌��ʐ��ɒ���
    ReDim Preserve result(LBound(sqlLines) To LBound(sqlLines) + resultCount - 1)
    
    RemovePLSQLComments = result
    
    Common.WriteLog "RemovePLSQLComments E"
End Function
