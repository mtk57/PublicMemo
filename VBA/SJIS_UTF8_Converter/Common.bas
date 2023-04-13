Attribute VB_Name = "Common"
Option Explicit

Public Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Public Declare PtrSafe Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

'���O�t�@�C���ԍ�
Private logfile_num As Integer
Private is_log_opened As Boolean

'-------------------------------------------------------------
'�w�肳�ꂽ�t�@�C����SJIS��UTF8(BOM����)�ϊ�����
' path : IN : �t�@�C���p�X(��΃p�X)
' is_backup : IN : True/False (True=�o�b�N�A�b�v����)
'                  ��������".bak_���ݓ���"��t�^
'-------------------------------------------------------------
Public Sub SJIStoUTF8(ByVal path As String, ByVal is_backup As Boolean)
    Dim in_str As String
    Dim buf As String
    Dim i As Integer
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS�`���̃e�L�X�g�t�@�C����ǂݍ���
    in_str = ""
    Open path For Input As #filenum
        '�e�L�X�g�����ׂĎ擾����
        Do Until EOF(filenum)
            Line Input #filenum, buf
            in_str = in_str & buf & vbLf
        Loop
    Close #filenum
        
    'Shift-JIS�ȊO�̃t�@�C����ǂݍ���ł��܂����ꍇ�͏I��
    For i = 1 To Len(in_str)
        If Asc(Mid(in_str, i, 1)) = -7295 Then Exit Sub
    Next
    
    '�o�b�N�A�b�v
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    'UTF-8�iBOM�t���j�Ńe�L�X�g�t�@�C���֏o��
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText in_str, 0
        .SaveToFile path, 2
        .Close
    End With
    
End Sub

'-------------------------------------------------------------
'�w�肳�ꂽ�t�@�C����UTF8(BOM����/�Ȃ�) �� SJIS�ϊ�����
' path : IN : �t�@�C���p�X(��΃p�X)
' is_backup : IN : True/False (True=�o�b�N�A�b�v����)
'                  ��������".bak_���ݓ���"��t�^
'-------------------------------------------------------------
Public Sub UTF8toSJIS(ByVal path As String, ByVal is_backup As Boolean)
    Dim in_str As String
    Dim out_str() As String
    Dim i As Integer
    
    'UTF-8��������UTF-8�iBOM�t���j�̃e�L�X�g�t�@�C����ǂݍ���
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8��������UTF-8�iBOM�t���j�ȊO��ǂݍ���ł��܂����ꍇ�͏I��
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                Exit Sub
            End If
        End If
    Next
    
    '���s���Ƀf�[�^�𕪂���
    out_str = Split(in_str, vbLf)
    
    '�o�b�N�A�b�v
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS�`���Ńe�L�X�g�t�@�C���֏o��
    Open path For Output As #filenum
        For i = 0 To UBound(out_str)
            Print #filenum, out_str(i)
        Next
    Close #filenum

End Sub

'-------------------------------------------------------------
'�t�@�C����UTF8(BOM����)���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=UTF8(BOM����), False=���L�ȊO)
'-------------------------------------------------------------
Public Function IsUTF8(ByVal path As String) As Boolean
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "�w�肳�ꂽ�t�@�C�������݂��܂��� (" & path & ")"
    End If

    Dim bytedata() As Byte: bytedata = ReadBinary(path, 3)
    Dim length As Integer: length = UBound(bytedata) + 1
    
    If length < 3 Then
        IsUTF8 = False
        Exit Function
    End If
    
    If bytedata(0) = &HEF And bytedata(1) = &HBB And bytedata(2) = &HBF Then
        IsUTF8 = True
    Else
        IsUTF8 = False
    End If
    
End Function

'-------------------------------------------------------------
'�t�@�C�����o�C�i���Ƃ��Ďw��T�C�Y�ǂݍ���
' path : IN : �t�@�C���p�X(��΃p�X)
' readsize : IN : �ǂݍ��ރT�C�Y
' Ret : �ǂݍ��񂾃o�C�i���z��
'-------------------------------------------------------------
Public Function ReadBinary(ByVal path As String, ByVal readsize As Integer) As Byte()
    Dim readdata() As Byte
    
    If readsize <= 0 Then
        ReadBinary = readdata()
        Exit Function
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    Open path For Binary Access Read As #filenum
    
    ReDim readdata(readsize - 1)
    
    Get #filenum, , readdata
    
    Close #filenum
    
    ReadBinary = readdata
End Function

'-------------------------------------------------------------
'�w��t�H���_�z���Ɏw��g���q�̃t�@�C�������݂��邩
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex. ".vb")
' Ret : True/False (True=���݂���, False=���݂��Ȃ�)
'-------------------------------------------------------------
Public Function IsExistsExtensionFile(ByVal path As String, ByVal ext As String) As Boolean
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim File As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    For Each subfolder In folder.subfolders
        If IsExistsExtensionFile(subfolder.path, ext) Then
            Set fso = Nothing
            Set folder = Nothing
            
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next subfolder
    
    For Each File In folder.Files
        If Right(File.Name, Len(ext)) = ext Then
            Set fso = Nothing
            Set folder = Nothing
        
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next File
    
    Set fso = Nothing
    Set folder = Nothing

    IsExistsExtensionFile = False
End Function

'-------------------------------------------------------------
'���O�t�@�C�����I�[�v������
' logfile_path : IN : ���O�t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub OpenLog(ByVal logfile_path As String)
    If is_log_opened = True Then
        '���łɃI�[�v�����Ă���̂Ŗ���
        Exit Sub
    End If
    logfile_num = FreeFile()
    Open logfile_path For Append As logfile_num
    is_log_opened = True
End Sub

'-------------------------------------------------------------
'���O�t�@�C���ɏ�������
' contents : IN : �������ޓ��e
'-------------------------------------------------------------
Public Sub WriteLog(ByVal contents As String)
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
End Sub

'-------------------------------------------------------------
'���O�t�@�C�����N���[�Y����
'-------------------------------------------------------------
Public Sub CloseLog()
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Close logfile_num
    logfile_num = -1
    is_log_opened = False
End Sub

'-------------------------------------------------------------
'�z��̋�s���폜����
' in_array : IN : ������z��
' Ret : ��s���폜�����z��
'-------------------------------------------------------------
Public Function DeleteEmptyArray(ByRef in_array() As String) As String()
    Dim ret_array() As String
    Dim i, cnt As Long
    Dim row As String
    
    ReDim ret_array(UBound(in_array))
    
    For i = LBound(in_array) To UBound(in_array)
        row = in_array(i)
        If Not IsEmpty(row) Then
            If row <> "" Then
                ret_array(cnt) = row
                cnt = cnt + 1
            End If
        End If
    Next
    
    ReDim Preserve ret_array(cnt - 1)
    
    DeleteEmptyArray = ret_array
End Function

'-------------------------------------------------------------
'�t�@�C�����X�g���쐬����
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g
'-------------------------------------------------------------
Public Function CreateFileList(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean) As String()
    Dim list() As String: list = CreateFileListMain(path, ext, is_subdir)
    CreateFileList = DeleteEmptyArray(list)
End Function

Private Function CreateFileListMain(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filelist() As String
    Dim cnt As Integer

    Dim File As String, f As Object
    File = Dir(path & "\" & ext)
    
    If File <> "" Then
        If Common.IsEmptyArray(filelist) = True Then
            cnt = 0
        Else
            cnt = UBound(filelist) + 1
        End If
        
        ReDim Preserve filelist(cnt)
        filelist(cnt) = path & "\" & File
    End If
    
    Do While File <> ""
        File = Dir()
        If File <> "" Then
            cnt = UBound(filelist) + 1
            ReDim Preserve filelist(cnt)
            filelist(cnt) = path & "\" & File
        End If
    Loop
    
    If is_subdir = False Then
        Set fso = Nothing
        CreateFileListMain = filelist
        Exit Function
    End If
    
    Dim filelist_sub() As String
    Dim filelist_merge() As String
    
    For Each f In fso.GetFolder(path).subfolders
        filelist_sub = CreateFileListMain(f.path, ext, is_subdir)
        filelist = Common.MergeArray(filelist_sub, filelist)
    Next f
    
    Set fso = Nothing
    CreateFileListMain = filelist
End Function

'-------------------------------------------------------------
'2�̔z����������ĕԂ�
' array1 : IN : �z��1
' array2 : IN : �z��2
' Ret : ���������z��
'-------------------------------------------------------------
Public Function MergeArray(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim merged As Variant
    merged = Split(Join(array1, vbCrLf) & vbCrLf & Join(array2, vbCrLf), vbCrLf)
    MergeArray = merged
End Function

'-------------------------------------------------------------
'2�̃e�L�X�g�t�@�C�����r���Ĉ�v���Ă��邩��Ԃ�
' file1 : IN : �t�@�C��1�p�X(��΃p�X)
' file2 : IN : �t�@�C��2�p�X(��΃p�X)
' Ret : ��r���� : True/False (True=��v)
'-------------------------------------------------------------
Public Function IsMatchTextFiles(ByVal file1 As String, ByVal file2 As String) As Boolean
    Dim filesize1 As Long: filesize1 = FileLen(file1)
    Dim filesize2 As Long: filesize2 = FileLen(file2)
    
    '�܂��t�@�C���T�C�Y�Ń`�F�b�N
    If filesize1 = 0 And filesize2 = 0 Then
        '�ǂ����0byte�Ȃ̂ň�v
        IsMatchTextFiles = True
        Exit Function
    ElseIf filesize1 <> filesize2 Then
        '�t�@�C���T�C�Y���قȂ�̂ŕs��v
        IsMatchTextFiles = False
        Exit Function
    ElseIf filesize1 = 0 Or filesize2 = 0 Then
        '�ǂ��炩��0byte�Ȃ̂ŕs��v
        IsMatchTextFiles = False
        Exit Function
    End If

    Dim fso1, fso2 As Object
    Dim ts1, ts2 As Object
    
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    
    Const READ_ONLY = 1
    Set ts1 = fso1.OpenTextFile(file1, READ_ONLY)
    Set ts2 = fso2.OpenTextFile(file2, READ_ONLY)
    
    Dim contents1 As String: contents1 = ts1.ReadAll
    Dim contents2 As String: contents2 = ts2.ReadAll
    
    ts1.Close
    ts2.Close
    Set ts1 = Nothing
    Set ts2 = Nothing
    Set fso1 = Nothing
    Set fso2 = Nothing
    
    IsMatchTextFiles = (contents1 = contents2)
End Function

'-------------------------------------------------------------
'������̔z��̖����ɕ������ǉ�����
' ary : IN/OUT : ������̔z��
' value : IN : �ǉ����镶����
'-------------------------------------------------------------
Public Sub AppendArray(ByRef ary() As String, ByVal value As String)
    Dim cnt As Integer: cnt = UBound(ary) + 1
    ReDim Preserve ary(cnt)
    ary(cnt) = value
End Sub

'-------------------------------------------------------------
'�t�H���_�p�X��񋓂���B�i�T�u�t�H���_�܂ށj
' ���ӁFpath�͖߂�l�ɂ͊܂܂Ȃ�
' path : IN : �t�H���_�p�X�i��΃p�X�j
' Ret : �t�H���_�p�X���X�g
'-------------------------------------------------------------
Public Function GetFolderPathList(ByVal path As String) As String()
    Dim fso As Object
    Dim top_dir As Object
    Dim sub_dir As Object
    Dim path_list() As String
    Dim dir_cnt As Long
    Dim i, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set top_dir = fso.GetFolder(path)

    dir_cnt = top_dir.subfolders.count
    If dir_cnt > 0 Then
        ReDim path_list(dir_cnt - 1)
        i = 0
        For Each sub_dir In top_dir.subfolders
            path_list(i) = sub_dir.path
            i = i + 1
            
            Dim sub_path_list() As String
            sub_path_list = GetFolderPathList(sub_dir.path)
            
            '�T�u�t�H���_���̃p�X��z��ɒǉ�����
            If sub_path_list(0) <> "" Then
                Dim cnt As Integer: cnt = UBound(path_list) + UBound(sub_path_list) + 1
                ReDim Preserve path_list(cnt)
                For j = LBound(sub_path_list) To UBound(sub_path_list)
                    path_list(i) = sub_path_list(j)
                    i = i + 1
                Next j
            End If
        Next sub_dir
        
        GetFolderPathList = path_list
    Else
        Dim ret_empty(0) As String
        GetFolderPathList = ret_empty
    End If
    
    Set sub_dir = Nothing
    Set top_dir = Nothing
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�H���_���R�s�[����(�T�u�t�H���_�܂�)
' src_path : IN : �R�s�[���t�H���_�p�X(��΃p�X)
' dst_path : IN : �R�s�[��t�H���_�p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub CopyFolder(ByVal src_path As String, dest_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�R�s�[���̃t�H���_�����݂��Ȃ��ꍇ�A�G���[�𔭐�������
    If Not fso.FolderExists(src_path) Then
        Err.Raise 53, , "�w�肳�ꂽ�t�H���_�����݂��܂���"
    End If
    
    '�R�s�[��̃t�H���_�����݂��Ȃ��ꍇ�A�쐬����
    If Not fso.FolderExists(dest_path) Then
        fso.CreateFolder dest_path
    End If
    
    '�R�s�[���̃t�H���_���̃t�@�C�����R�s�[����
    Dim File As Object
    For Each File In fso.GetFolder(src_path).Files
        fso.CopyFile File.path, fso.BuildPath(dest_path, File.Name), True
    Next
    
    '�R�s�[���̃t�H���_���̃T�u�t�H���_���R�s�[����
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(src_path).subfolders
        CopyFolder subfolder.path, fso.BuildPath(dest_path, subfolder.Name)
    Next
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'Yes/No���b�Z�[�W�{�b�N�X��\������
' msg : IN : ���b�Z�[�W
' Ret : True/False (True=Yes, False=No)
'-------------------------------------------------------------
Public Function ShowYesNoMessageBox(ByVal msg As String) As Boolean
    Dim result As Integer: result = MsgBox(msg, vbYesNo, "Confirm")
    
    If result = vbYes Then
        ShowYesNoMessageBox = True
    Else
        ShowYesNoMessageBox = False
    End If
End Function

'-------------------------------------------------------------
'�O���A�v���P�[�V���������s���A�I������܂őҋ@����
' exe_path : IN : �O���A�v���P�[�V����(exe)�̐�΃p�X
'                 exe�ɓn���p�����[�^������ꍇ���ꏏ�ɏ�������
' Ret : �v���Z�X�̖߂�l
'-------------------------------------------------------------
Public Function RunProcessWait(ByVal exe_path As String) As Long

  Dim wsh As Object
  Set wsh = CreateObject("Wscript.Shell")
  
  Const NOT_DISP = 0
  Const DISP = 1
  Const WAIT = True
  Const NO_WAIT = False
  
  Dim Process As Object
  Set Process = wsh.Exec(exe_path)

  '�v���Z�X�������ɒʒm���󂯎��
  Do While Process.Status = 0
    DoEvents
  Loop

  '�v���Z�X�̖߂�l���擾����
  RunProcessWait = Process.ExitCode

  Set Process = Nothing
  Set wsh = Nothing
End Function

'-------------------------------------------------------------
'�p�X������̖�����\���������ĕԂ�
' path : IN : �p�X������
' Ret : �p�X������
'-------------------------------------------------------------
Public Function RemoveTrailingBackslash(ByVal path As String) As String
    If Right(path, 1) = "\" Then
        path = Left(path, Len(path) - 1)
    End If
    RemoveTrailingBackslash = path
End Function

'-------------------------------------------------------------
'�t�@�C���̓��e���w�肳�ꂽ�V�[�g�ɏo�͂���
' file_path : IN : �t�@�C���p�X (��΃p�X)
' sheet_name : IN : �V�[�g��
' is_sjis : IN :�����t�@�C���̃G���R�[�h�w��BTrue/False (True=Shift-JIS, False=UTF-16)  TODO:�����ꎩ�����ʂ��������B�B�B
'-------------------------------------------------------------
Public Sub OutputTextFileToSheet(ByVal file_path As String, ByVal sheet_name As String, ByVal is_sjis As Boolean)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�@�C�����J��
    Dim file_format As Integer
    Const FORMAT_ASCII = 0
    Const FORMAT_UNICODE = -1
    
    If is_sjis = True Then
        file_format = FORMAT_ASCII
    Else
        file_format = FORMAT_UNICODE
    End If
    
    Dim File As Object
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    
    Set File = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, file_format)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    '�t�@�C���̓��e���V�[�g�ɏo��
    Dim row As Integer: row = 1
    
    Do While Not File.AtEndOfStream
        ws.Cells(row, 1).value = File.ReadLine
        row = row + 1
    Loop
    
    File.Close
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'SJIS�Ńe�L�X�g�t�@�C�����쐬����
' contents : IN : ���e
' path : IN : �t�@�C���p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub CreateSJISTextFile(ByRef contents() As String, ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim txt As Object
    Dim i As Integer
    
    Dim IS_OVERWRITE As Boolean: IS_OVERWRITE = True
    Dim IS_UNICODE As Boolean: IS_UNICODE = False
    
    Set txt = fso.CreateTextFile(path, IS_OVERWRITE, IS_UNICODE)
    
    For i = LBound(contents) To UBound(contents)
        txt.WriteLine contents(i)
    Next i
    
    txt.Close
    Set fso = Nothing
End Sub


'-------------------------------------------------------------
'�T�u�t�H���_���܂Ƃ߂č쐬����
' path : IN : �t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub CreateFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folders() As String
    folders = Split(path, Application.PathSeparator)
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(folders) To UBound(folders)
        ary = folders
        ReDim Preserve ary(i)
        If Not fso.FolderExists(Join(ary, Application.PathSeparator)) Then
            Call fso.CreateFolder(Join(ary, Application.PathSeparator))
        End If
    Next
  
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�H���_���폜����
' path : IN : �t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub DeleteFolder(ByVal path As String)
    If IsExistsFolder(path) = False Then
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.DeleteFolder path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'������z��̋��ʕ������Ԃ�
' list : IN : ������z��
' Ret : ���ʕ�����
'       Ex. list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'           Ret = "hoge"
'-------------------------------------------------------------
Function GetCommonString(ByRef list() As String) As String
    Dim common_string As String
    Dim i, j As Integer
    Dim flag As Boolean
    
    '�ŏ��̕���������ʕ�����̏����l�Ƃ���
    common_string = list(0)
    
    '�e��������r����
    For i = 1 To UBound(list)
        flag = False
        '���ʕ������擾����
        For j = 1 To Len(common_string)
            If Mid(common_string, j, 1) <> Mid(list(i), j, 1) Then
                common_string = Left(common_string, j - 1)
                flag = True
                Exit For
            End If
        Next j
    Next i
    
    '���ʂ��o�͂���
    GetCommonString = common_string
End Function

'-------------------------------------------------------------
'��΃t�@�C���p�X�̐e�t�H���_�p�X���擾����
' path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �e�t�H���_�p�X (��΃p�X)
'       Ex. path = "C:\tmp\abc.txt"
'           Ret = "C:\tmp"
'-------------------------------------------------------------
Public Function GetFolderNameFromPath(ByVal path As String) As String
    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = Left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'-------------------------------------------------------------
'���΃p�X���΃p�X�ɕϊ�����
' base_path : IN : ��ƂȂ�t�H���_�p�X(��΃p�X)
' ref_path : IN : �t�@�C���p�X�i���΃p�X)
' Ret : ��΃p�X
'       Ex. base_path = "C:\tmp\abc"
'           ref_path = "..\cdf\xyz.txt"
'           Ret = "C:\tmp\cdf\xyz.txt"
'-------------------------------------------------------------
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C���̑��݃`�F�b�N���s��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistsFile(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        IsExistsFile = True
    Else
        IsExistsFile = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�H���_�̑��݃`�F�b�N���s��
' path : IN : �t�H���_�p�X(��΃p�X)
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistsFolder(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C��������g���q��Ԃ�
' filename : IN : �t�@�C����
' Ret : �t�@�C�����̊g���q
'        Ex. "abc.txt"�̏ꍇ�A"txt"���Ԃ�
'            "."���܂܂�Ă��Ȃ��ꍇ��""���Ԃ�
'-------------------------------------------------------------
Public Function GetFileExtension(ByVal filename As String) As String
    Dim dot_pos As Integer
    
    ' "."�̈ʒu���擾
    dot_pos = InStrRev(filename, ".")
    
    ' �g���q���擾
    If dot_pos > 0 Then
        GetFileExtension = Right(filename, Len(filename) - dot_pos)
    Else
        GetFileExtension = ""
    End If
End Function

'-------------------------------------------------------------
'�w��t�H���_�z�����w��t�@�C�����Ō������Ă��̓��e��Ԃ�
' target_folder : IN :�����t�H���_�p�X(��΃p�X)
' target_file : IN :�����t�@�C����
' is_sjis : IN :�����t�@�C���̃G���R�[�h�w��BTrue/False (True=Shift-JIS, False=UTF-8)  TODO:�����ꎩ�����ʂ��������B�B�B
' Ret : �ǂݍ��񂾃t�@�C���̓��e
'       �z��̖����ɂ͌����t�@�C���̐�΃p�X���i�[����
'-------------------------------------------------------------
Public Function SearchAndReadFiles(ByVal target_folder As String, ByVal target_file As String, ByVal is_sjis As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim File As Object
    For Each File In folder.Files
        If fso.FileExists(File.path) And fso.GetFileName(File.path) Like target_file Then
            '�����Ώۂ̃t�@�C����ǂݍ���
            Dim contents As String
            
            If is_sjis = True Then
                'S-JIS
                contents = ReadTextFileBySJIS(File.path)
            Else
                'UTF-8
                contents = ReadTextFileByUTF8(File.path)
            End If
            
            '�t�@�C���̓��e��z��Ɋi�[����
            Dim lines() As String: lines = Split(contents, vbCrLf)
            
            '�����Ƀt�@�C���p�X��ǉ�����
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = File.path
            SearchAndReadFiles = lines
            Set fso = Nothing
            Exit Function
        End If
    Next File
    
    '�T�u�t�H���_����������
    Dim subfolder As Object
    For Each subfolder In folder.subfolders
        Dim result() As String
        result = SearchAndReadFiles(subfolder.path, target_file, is_sjis)
        If UBound(result) >= 1 Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ�́A��̔z���Ԃ�
    Dim ret_empty(0) As String
    SearchAndReadFiles = ret_empty
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'SJIS�`���̃e�L�X�g�t�@�C����ǂݍ���
' file_path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function ReadTextFileBySJIS(ByVal file_path) As String
    'TODO:�����`�F�b�N
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const FORMAT_ASCII = 0
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    Dim contents As String
    
    Dim File As Object
    Set File = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    contents = File.ReadAll
    
    File.Close
    Set File = Nothing
    Set fso = Nothing
    
    ReadTextFileBySJIS = contents
End Function

'-------------------------------------------------------------
'UTF-8�`���̃e�L�X�g�t�@�C����ǂݍ���
' file_path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function ReadTextFileByUTF8(ByVal file_path) As String
    'TODO:�����`�F�b�N
    
    Dim contents As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file_path
        contents = .ReadText
        .Close
    End With
    
    ReadTextFileByUTF8 = contents
End Function

'-------------------------------------------------------------
'�z�񂪋󂩂��`�F�b�N����
' arg : IN : �z��
' Ret : True/False (True=��)
'-------------------------------------------------------------
Public Function IsEmptyArray(arg As Variant) As Boolean
    On Error Resume Next
    IsEmptyArray = Not (UBound(arg) > 0)
    IsEmptyArray = CBool(Err.Number <> 0)
End Function

'-------------------------------------------------------------
'���ݓ����𕶎���ŕԂ�
' Ret :Ex."20230326123456"
'-------------------------------------------------------------
Public Function GetNowTimeString() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString = str_date & str_time
End Function

'-------------------------------------------------------------
'�V�[�g�̑��݃`�F�b�N
' sheet_name : IN : �V�[�g��
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistSheet(ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name = sheet_name Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'-------------------------------------------------------------
'�V�[�g��ǉ�����
' sheet_name : IN : �V�[�g��
'-------------------------------------------------------------
Public Sub AddSheet(ByVal sheet_name As String)
    If IsExistSheet(sheet_name) = True Then
        Application.DisplayAlerts = False
        Sheets(sheet_name).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheet_name
End Sub

