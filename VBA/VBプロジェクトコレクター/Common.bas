Attribute VB_Name = "Common"
Option Explicit

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
    Dim FORMAT_ASCII As Integer: FORMAT_ASCII = 0
    Dim FORMAT_UNICODE As Integer: FORMAT_UNICODE = -1
    
    If is_sjis = True Then
        file_format = FORMAT_ASCII
    Else
        file_format = FORMAT_UNICODE
    End If
    
    Dim ts As Object
    Dim READ_ONLY As Integer: READ_ONLY = 1
    Dim IS_CREATE_FILE As Boolean: IS_CREATE_FILE = False
    
    Set ts = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, file_format)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    '�t�@�C���̓��e���V�[�g�ɏo��
    Dim row As Integer: row = 1
    
    Do While Not ts.AtEndOfStream
        ws.Cells(row, 1).value = ts.ReadLine
        row = row + 1
    Loop
    
    ts.Close
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
    
    Dim file As Object
    For Each file In folder.Files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like target_file Then
            '�����Ώۂ̃t�@�C����ǂݍ���
            Dim contents As String
            
            If is_sjis = True Then
                'S-JIS
                contents = ReadTextFileBySJIS(file.path)
            Else
                'UTF-8
                contents = ReadTextFileByUTF8(file.path)
            End If
            
            '�t�@�C���̓��e��z��Ɋi�[����
            Dim lines() As String: lines = Split(contents, vbCrLf)
            
            '�����Ƀt�@�C���p�X��ǉ�����
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = file.path
            SearchAndReadFiles = lines
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    '�T�u�t�H���_����������
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        Dim result() As String
        result = SearchAndReadFiles(subFolder.path, target_file, is_sjis)
        If IsEmptyArray(result) = False Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subFolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ�́A��̔z���Ԃ�
    Dim ret_empty() As String
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
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    contents = ts.ReadAll
    ts.Close
    Set ts = Nothing
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

