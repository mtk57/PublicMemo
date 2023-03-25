Attribute VB_Name = "Common"
Option Explicit

'�T�u�t�H���_���܂Ƃ߂č쐬����
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

'������z��̋��ʕ������Ԃ�
'Ex.
'  list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'  Ret = "hoge"
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
        
        '���ʕ������Ȃ��ꍇ�̓��[�v���I������
        If flag = False Then
            Exit For
        End If
    Next i
    
    '���ʂ��o�͂���
    GetCommonString = common_string
End Function

'��΃t�@�C���p�X�̐e�t�H���_�p�X���擾����
'Ex.  path=C:\tmp\abc.txt
'     Ret=C:\tmp
Public Function GetFolderNameFromPath(ByVal path As String) As String
    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = Left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'���΃p�X���΃p�X�ɕϊ�����
'Ex.  base_path=C:\tmp\abc
'     ref_path=..\cdf\xyz.txt
'     Ret=C:\tmp\cdf\xyz.txt
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'�t�H���_�̑��݃`�F�b�N���s��
'path�͐�΃p�X�Ƃ���
Public Function IsExistsFolder(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'�t�@�C��������g���q��Ԃ�
'Ex. "abc.txt"�̏ꍇ�A"txt"���Ԃ�
'"."���܂܂�Ă��Ȃ��ꍇ��""���Ԃ�
Public Function GetFileExtension(filename As String) As String
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


'�w��t�H���_�z�����w��t�@�C�����Ō������Ă��̓��e��Ԃ�
'�ǂݍ��񂾃t�@�C���̓��e��1�s����String�z��ƂȂ邪�A
'�z��̖����ɂ̓t�@�C���̐�΃p�X���i�[����̂Œ��ӁB
Public Function SearchAndReadFiles(target_folder As String, target_file As String, is_sjis As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim file As Object
    For Each file In folder.Files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like target_file Then
            '�����Ώۂ̃t�@�C����ǂݍ���
            Dim ts As Object
            If is_sjis = True Then
                Set ts = fso.OpenTextFile(file.path, 1, False, 0)
            Else
                Set ts = fso.OpenTextFile(file.path, 1, False, 1)
            End If
            Dim fileContent As String
            fileContent = ts.ReadAll
            ts.Close
            
            '�t�@�C���̓��e��z��Ɋi�[���ĕԂ�
            Dim lines() As String
            lines = Split(fileContent, vbCrLf)
            
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
        If Not IsEmpty(result) Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subFolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ�́A��̔z���Ԃ�
    SearchAndReadFiles = Split("", vbCrLf)
    Set fso = Nothing
End Function


'���ݓ����𕶎���ŕԂ�
Public Function GetNowTimeString() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString = str_date & str_time
End Function

'�V�[�g�̑��݃`�F�b�N
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

'�V�[�g��ǉ�����
Public Sub AddSheet(ByVal sheet_name As String)
    If IsExistSheet(sheet_name) = True Then
        Application.DisplayAlerts = False
        Sheets(sheet_name).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheet_name
End Sub

