Attribute VB_Name = "Common"
Option Explicit

Public Const VERSION = "1.0.6"

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
'�t�@�C�����R�s�[����
' src_path : I : �R�s�[���t�@�C���p�X(��΃p�X)
' dst_path : I : �R�s�[��t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub CopyFile(ByVal src_path As String, ByVal dst_path As String)
    If IsExistsFile(src_path) = False Then
        Err.Raise 53, , "[CopyFile] �w�肳�ꂽ�t�@�C�������݂��܂��� (src_path=" & src_path & ")"
    End If
    
    If dst_path = "" Or src_path = dst_path Or IsExistsFile(dst_path) = True Then
        Exit Sub
    End If
    
    FileCopy src_path, dst_path
End Sub


'-------------------------------------------------------------
'�t�H���_�����l�[������
' path : I : �t�H���_�p�X(��΃p�X)
' rename : I : ���l�[����̃t�H���_��
' Ret : ���l�[����̃t�H���_�p�X
'-------------------------------------------------------------
Public Function RenameFolder(ByVal path As String, ByVal rename As String) As String
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[RenameFolder] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(path)
    
    folder.name = rename
    RenameFolder = folder.path
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�w���̑S�s���w�胏�[�h�Ō������A�q�b�g�����s�ԍ���Ԃ�
' ws : I : ���[�N�V�[�g
' find_clm : I : �w���(Ex."A")
' find_start_row : I : �����J�n�s(1�n�܂�)
' keyword : I : �������[�h
' Ret : �q�b�g�����s�ԍ�
'-------------------------------------------------------------
Public Function FindRowByKeyword( _
  ByVal ws As Worksheet, _
  ByVal find_clm As String, _
  ByVal find_start_row As Long, _
  ByVal keyword As String _
) As Long
    Dim rng As Range
    Dim cell As Range
    Dim found_row As Long
    
    Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlUp).row)
    
    found_row = 0
    For Each cell In rng
        If cell.value = keyword Then
            found_row = cell.row
            Exit For
        End If
    Next cell
    
    FindRowByKeyword = found_row
End Function

'-------------------------------------------------------------
'�V�[�g�̓��e��2�����z��Ɋi�[����
' sheet_name : I : �V�[�g��
' Ret : �V�[�g�̓��e
'-------------------------------------------------------------
Public Function GetSheetContentsByStringArray(ByVal sheet_name As String) As String()
    Dim ws As Worksheet
    Dim arr() As String
    Dim row_cnt As Long, clm_cnt As Long
    Dim r As Long, c As Long
    
    Set ws = ActiveWorkbook.Worksheets(sheet_name)

    row_cnt = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    clm_cnt = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ReDim arr(1 To row_cnt, 1 To clm_cnt)
    For r = 1 To row_cnt
        For c = 1 To clm_cnt
            arr(r, c) = CStr(ws.Cells(r, c).value)
        Next c
    Next r

    GetSheetContentsByStringArray = arr
End Function

'-------------------------------------------------------------
'�g���q��ύX����
' path : I : �t�@�C���p�X(��΃p�X)
' ext : I : �ύX��̊g���q(Ex. ".new")
' Ret : �ύX��̃t�@�C���p�X(��΃p�X)
'       path�̃t�@�C�������݂��Ȃ��ꍇ��path��Ԃ�
'-------------------------------------------------------------
Public Function ChangeFileExt(ByVal path As String, ByVal ext As String) As String
    If Common.IsExistsFile(path) = False Then
        'Err.Raise 53, , "[ChangeFileExt] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
        ChangeFileExt = path
        Exit Function
    End If
    
    Dim SEP As String: SEP = Application.PathSeparator
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim old_ext As String: old_ext = fso.GetExtensionName(path)
    Dim file_name As String: file_name = fso.GetBaseName(path)
    Dim new_path As String
    
    '�V�����g���q�ɕύX
    file_name = file_name & ext
    new_path = fso.GetParentFolderName(path) & SEP & file_name
    
    '�t�@�C������ύX
    fso.MoveFile path, new_path
    Set fso = Nothing
    
    ChangeFileExt = new_path
End Function

'-------------------------------------------------------------
'�u�b�N���J���ăV�[�g���擾����
' book_path : IN : Excel�t�@�C���p�X(��΃p�X)
' sheet_name : IN : �V�[�g��
' visible : IN : True/False (True=�\��, False=��\��)
' Ret : �V�[�g�I�u�W�F�N�g
'-------------------------------------------------------------
Public Function GetSheet(ByVal book_path As String, ByVal sheet_name As String, ByVal visible As Boolean) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filename:=book_path, UpdateLinks:=False, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    ActiveWindow.visible = visible
    Set GetSheet = wb.Worksheets(sheet_name)
End Function

'-------------------------------------------------------------
'�u�b�N�ƕ���
' name : IN : �u�b�N��(Excel�t�@�C����)
'-------------------------------------------------------------
Public Sub CloseBook(ByVal name As String)
    Dim wb As Workbook
    For Each wb In Workbooks
        If InStr(wb.name, name) > 0 Then
            wb.Close SaveChanges:=False
        End If
    Next
End Sub

'-------------------------------------------------------------
'�t�@�C�����폜����
' path : IN : �t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub DeleteFile(ByVal path As String)
    If Common.IsExistsFile(path) = False Then
        Exit Sub
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const DELETE_READONLY = True
    fso.DeleteFile path, DELETE_READONLY
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�@�C���������j�[�N�Ȗ��̂Ƀ��l�[�����ăR�s�[����
' src_file_path : IN : �R�s�[���t�@�C���p�X(��΃p�X)
' dst_dir_path : IN : �R�s�[��t�H���_�p�X(��΃p�X)
'                     ������\�͕s�v
'                     ��̏ꍇ�̓R�s�[���Ɠ����t�H���_�Ƃ���
' Ret : ���l�[���R�s�[��̃t�@�C���p�X
'-------------------------------------------------------------
Public Function CopyUniqueFile(ByVal src_file_path As String, ByVal dst_dir_path As String) As String
    If Common.IsExistsFile(src_file_path) = False Then
        CopyUniqueFile = ""
        Exit Function
    End If
    
    Dim SEP As String: SEP = Application.PathSeparator
    Dim dst_file_path As String
    
    Dim unique_filename As String: unique_filename = Common.GetFileName(src_file_path) & ".bak_" & GetNowTimeString()
    
    If dst_dir_path = "" Then
        dst_file_path = Common.GetFolderNameFromPath(src_file_path) & SEP & unique_filename
    Else
        dst_file_path = dst_dir_path & SEP & unique_filename
    End If

    FileCopy src_file_path, dst_file_path
    
    CopyUniqueFile = dst_file_path
End Function

'-------------------------------------------------------------
'�t�@�C������Ԃ�
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : �t�@�C����
'-------------------------------------------------------------
Public Function GetFileName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�w��t�H���_�z�����w��t�@�C�����Ō������ăt�@�C���p�X��Ԃ�
' search_path : IN : �����t�H���_�p�X(��΃p�X)
' search_name : IN : �����t�@�C����
' Ret : �t�@�C���p�X
'-------------------------------------------------------------
Public Function SearchFile(ByVal search_path As String, ByVal search_name As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(search_path)
    
    Dim file As Object
    For Each file In folder.files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like search_name Then
            '����
            SearchFile = file.path
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    '�T�u�t�H���_����������
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        Dim result As String
        result = SearchFile(subfolder.path, search_name)
        If result <> "" Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchFile = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ
    SearchFile = ""
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�w��t�H���_��UTF8��S��SJIS�ɂ���
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g
'-------------------------------------------------------------
Public Sub UTF8toSJIS_AllFile(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean)
    If Common.IsExistsFolder(path) = False Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If
    
    If ext = "" Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] �g���q���w�肳��Ă��܂���"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(path, ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        Common.UTF8toSJIS src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'�w��t�H���_��SJIS��S��UTF8�ɂ���
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g
'-------------------------------------------------------------
Public Sub SJIStoUTF8_AllFile(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean)
    If Common.IsExistsFolder(path) = False Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If
    
    If ext = "" Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] �g���q���w�肳��Ă��܂���"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(path, ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        Common.SJIStoUTF8 src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'�w�肳�ꂽ�t�@�C����SJIS��UTF8(BOM����)�ϊ�����
' path : IN : �t�@�C���p�X(��΃p�X)
' is_backup : IN : True/False (True=�o�b�N�A�b�v����)
'                  ��������".bak_���ݓ���"��t�^
'-------------------------------------------------------------
Public Sub SJIStoUTF8(ByVal path As String, ByVal is_backup As Boolean)
    Dim in_str As String
    Dim buf As String
    Dim i As Long
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS�`���̃e�L�X�g�t�@�C����ǂݍ���
    in_str = ""
    Open path For Input As #filenum
        '�e�L�X�g�����ׂĎ擾����
        Do Until EOF(filenum)
            Line Input #filenum, buf
            in_str = in_str & buf & vbCrLf
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
    Dim i As Long
    
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
    out_str = Split(in_str, vbCrLf)
    
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
'�t�@�C����SJIS���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=SJIS)
'-------------------------------------------------------------
Public Function IsSJIS(ByVal path As String) As Boolean
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsSJIS] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If
    
    Dim Ado As Object
    Const TYPE_BINARY = 1
    Set Ado = CreateObject("ADODB.Stream")
    Ado.Type = TYPE_BINARY
    Ado.Open

    Ado.LoadFromFile path
    Dim read_data As String: read_data = Ado.Read
    Ado.Close
    Set Ado = Nothing

    Dim i As Long
    Dim first_byte As Byte
    Dim second_byte As Byte
    Dim is_dbcs As Boolean
    
    For i = 1 To LenB(read_data)

        first_byte = AscB(MidB(read_data, i, 1))

        '�S�p������(DBCS)�̐擪1�o�C�g�ł��邩
        is_dbcs = False

        If &H81 <= first_byte And first_byte <= &H9F Then
            is_dbcs = True
        ElseIf &HE0 <= first_byte And first_byte <= &HEF Then
            is_dbcs = True
        End If

        If is_dbcs Then
            i = i + 1

            If i > LenB(read_data) Then
                IsSJIS = False
                Exit Function
            End If

            second_byte = AscB(MidB(read_data, i, 1))

            If &H40 <= second_byte And second_byte <= &H7F Then
                'SJIS!
            ElseIf &H80 <= second_byte And second_byte <= &HFC Then
                'SJIS!
            Else
                IsSJIS = False
                Exit Function
            End If
        End If
    Next

    IsSJIS = True
End Function

'-------------------------------------------------------------
'�t�@�C����UTF8(BOM����/�Ȃ�)���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=UTF8(BOM����/�Ȃ�))
'-------------------------------------------------------------
Public Function IsUTF8(ByVal path As String) As Boolean
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If
    
    Dim in_str As String
    Dim out_str() As String
    Dim i As Long
    
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
                IsUTF8 = False
                Exit Function
            End If
        End If
    Next
    
    IsUTF8 = True
End Function

'-------------------------------------------------------------
'�t�@�C����UTF8(BOM����)���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=UTF8(BOM����), False=���L�ȊO)
'-------------------------------------------------------------
Public Function IsUTF8_WithBom(ByVal path As String) As Boolean
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8_WithBom] �w�肳�ꂽ�t�@�C�������݂��܂��� (path" & path & ")"
    End If

    Dim bytedata() As Byte: bytedata = ReadBinary(path, 3)
    Dim length As Integer: length = UBound(bytedata) + 1
    
    If length < 3 Then
        IsUTF8_WithBom = False
        Exit Function
    End If
    
    If bytedata(0) = &HEF And bytedata(1) = &HBB And bytedata(2) = &HBF Then
        IsUTF8_WithBom = True
    Else
        IsUTF8_WithBom = False
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
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    For Each subfolder In folder.SubFolders
        If IsExistsExtensionFile(subfolder.path, ext) Then
            Set fso = Nothing
            Set folder = Nothing
            
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next subfolder
    
    For Each file In folder.files
        If Right(file.name, Len(ext)) = ext Then
            Set fso = Nothing
            Set folder = Nothing
        
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next file
    
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
    
    cnt = 0
    
    If IsEmptyArray(in_array) = True Then
        GoTo FINISH
    End If
    
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
    
FINISH:
    If cnt > 0 Then
        ReDim Preserve ret_array(cnt - 1)
    Else
        ReDim ret_array(0)
    End If
    
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

    Dim file As String, f As Object
    file = Dir(path & "\" & ext)
    
    If file <> "" Then
        If Common.IsEmptyArray(filelist) = True Then
            cnt = 0
        Else
            cnt = UBound(filelist) + 1
        End If
        
        ReDim Preserve filelist(cnt)
        filelist(cnt) = path & "\" & file
    End If
    
    Do While file <> ""
        file = Dir()
        If file <> "" Then
            cnt = UBound(filelist) + 1
            ReDim Preserve filelist(cnt)
            filelist(cnt) = path & "\" & file
        End If
    Loop
    
    If is_subdir = False Then
        Set fso = Nothing
        CreateFileListMain = filelist
        Exit Function
    End If
    
    Dim filelist_sub() As String
    Dim filelist_merge() As String
    
    For Each f In fso.GetFolder(path).SubFolders
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
    
    'TODO:�o�C�i�����x���Ŕ�r���ׂ�
    
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

    dir_cnt = top_dir.SubFolders.count
    If dir_cnt > 0 Then
        ReDim path_list(dir_cnt - 1)
        i = 0
        For Each sub_dir In top_dir.SubFolders
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
        Err.Raise 53, , "[CopyFolder] �w�肳�ꂽ�t�H���_�����݂��܂���B(src_path=" & src_path & ")"
    End If
    
    '�R�s�[��̃t�H���_�����݂��Ȃ��ꍇ�A�쐬����
    If Not fso.FolderExists(dest_path) Then
        fso.CreateFolder dest_path
    End If
    
    '�R�s�[���̃t�H���_���̃t�@�C�����R�s�[����
    Dim file As Object
    For Each file In fso.GetFolder(src_path).files
        fso.CopyFile file.path, fso.BuildPath(dest_path, file.name), True
    Next
    
    '�R�s�[���̃t�H���_���̃T�u�t�H���_���R�s�[����
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(src_path).SubFolders
        CopyFolder subfolder.path, fso.BuildPath(dest_path, subfolder.name)
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
'-------------------------------------------------------------
Public Sub OutputTextFileToSheet(ByVal file_path As String, ByVal sheet_name As String)
    If Common.IsExistsFile(file_path) = False Or sheet_name = "" Then
        Err.Raise 53, , "[OutputTextFileToSheet] �w�肳�ꂽ�t�@�C�������݂��܂��� (file_path=" & file_path & ")"
    End If

    '���[�N�p�ɃR�s�[����
    Dim wk As String: wk = Common.CopyUniqueFile(file_path, "")
    
    '���[�N�t�@�C����SJIS�ɕϊ�����
    Common.UTF8toSJIS wk, False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�@�C�����J��
    Const FORMAT_ASCII = 0
    Const FORMAT_UNICODE = -1
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    '�t�@�C���̓��e���V�[�g�ɏo��
    Dim row As Integer: row = 1
    
    Do While Not fileobj.AtEndOfStream
        ws.Cells(row, 1).value = fileobj.ReadLine
        row = row + 1
    Loop
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    '���[�N�t�@�C�����폜����
    DeleteFile wk
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
    Dim i As Long
    
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
'�t�H���_���ړ�����
' src_path : IN : �ړ����t�H���_�p�X (��΃p�X)
' dst_path : IN : �ړ���t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub MoveFolder(ByVal src_path As String, ByVal dst_path As String)
    If IsExistsFolder(src_path) = False Then
        Err.Raise 53, , "[MoveFolder] �ړ����t�H���_�����݂��܂��� (src_path=" & src_path & ")"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.MoveFolder src_path, dst_path
    
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
    Dim i, j As Long
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
        GetFileExtension = LCase(Right(filename, Len(filename) - dot_pos))
    Else
        GetFileExtension = ""
    End If
End Function

'-------------------------------------------------------------
'�w��t�H���_�z�����w��t�@�C�����Ō������Ă��̓��e��Ԃ�
' target_folder : IN :�����t�H���_�p�X(��΃p�X)
' target_file : IN :�����t�@�C����
' Ret : �ǂݍ��񂾃t�@�C���̓��e
'       �z��̖����ɂ͌����t�@�C���̐�΃p�X���i�[����
'-------------------------------------------------------------
Public Function SearchAndReadFiles(ByVal target_folder As String, ByVal target_file As String) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim fileobj As Object
    For Each fileobj In folder.files
        If fso.FileExists(fileobj.path) And fso.GetFileName(fileobj.path) Like target_file Then
            '�����Ώۂ̃t�@�C����ǂݍ���
            Dim contents As String: contents = ReadTextFileBySJIS(fileobj.path)

            '�t�@�C���̓��e��z��Ɋi�[����
            Dim lines() As String: lines = Split(contents, vbCrLf)
            
            '�����Ƀt�@�C���p�X��ǉ�����
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = file.path
            SearchAndReadFiles = lines
            Set fileobj = Nothing
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    '�T�u�t�H���_����������
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
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
'SJIS�Ńe�L�X�g�t�@�C����ǂݍ���
'��UTF8�̃t�@�C����SJIS�ɕϊ����ēǂݍ���!
' path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function ReadTextFileBySJIS(ByVal path As String) As String
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "[ReadTextFileBySJIS] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If
    
    '���[�N�p�ɃR�s�[����
    Dim wk As String: wk = Common.CopyUniqueFile(path, "")
    
    '���[�N�t�@�C����SJIS�ɕϊ�����
    Common.UTF8toSJIS wk, False
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const FORMAT_ASCII = 0
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    Dim contents As String: contents = fileobj.ReadAll
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    '���[�N�t�@�C�����폜����
    DeleteFile wk
    
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
        If ws.name = sheet_name Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'-------------------------------------------------------------
'�V�[�g���폜����
' sheet_name : IN : �V�[�g��
'-------------------------------------------------------------
Public Sub DeleteSheet(ByVal sheet_name As String)
    If IsExistSheet(sheet_name) = False Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Sheets(sheet_name).Delete
    Application.DisplayAlerts = True
End Sub

'-------------------------------------------------------------
'�V�[�g��ǉ�����
' sheet_name : IN : �V�[�g��
'-------------------------------------------------------------
Public Sub AddSheet(ByVal sheet_name As String)
    DeleteSheet sheet_name
    Worksheets.Add.name = sheet_name
End Sub

'-------------------------------------------------------------
'�u�b�N���A�N�e�B�u�ɂ���
' book_name : IN : �u�b�N��(Excel�t�@�C����)
'-------------------------------------------------------------
Public Sub ActiveBook(ByVal book_name As String)
    Dim wb As Workbook
    Set wb = Workbooks(book_name)
    wb.Activate
End Sub



