Attribute VB_Name = "Process"
Option Explicit

'�萔
Private Const MAIN_SHEET = "main"
Private Const SEARCH_FILE_NAME = "O5"
Private Const SEARCH_DIR_PATH = "O6"
Private Const FILE_ENCODE = "O7"
Private Const OUT_DIR_PATH = "O11"
Private Const OUT_SHEET_NAME = "O12"
Private Const OUT_BAT_PATH = "O13"

'�V�[�g������W�������
Private search_file As String
Private search_path As String
Private encode As String
Private out_path As String
Private out_sheet As String
Private out_bat As String

'���C������
Public Sub Run()
    Worksheets(MAIN_SHEET).Activate
    Dim err_msg As String

    'main�V�[�g�̏������W
    search_file = Range(SEARCH_FILE_NAME).value
    search_path = Range(SEARCH_DIR_PATH).value
    encode = Range(FILE_ENCODE).value
    out_path = Range(OUT_DIR_PATH).value
    out_sheet = Range(OUT_SHEET_NAME).value
    out_bat = Range(OUT_BAT_PATH).value
    
    '���W�����������؂���
    err_msg = Validate()
    If err_msg <> "" Then
        MsgBox err_msg
        Exit Sub
    End If
    
    '�t�@�C���G���R�[�h
    Dim is_sjis As Boolean: is_sjis = True
    If encode = "UTF-8" Then
        is_sjis = False
    End If
    
    'VB�v���W�F�N�g�t�@�C�����������ēǂݍ���
    Dim contents() As String: contents = Common.SearchAndReadFiles(search_path, search_file, is_sjis)
    
    If UBound(contents) = -1 Then
        MsgBox "VB�v���W�F�N�g�t�@�C����������܂���ł���"
        Exit Sub
    End If
    
    'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
    Dim filelist() As String: filelist = ParseContents(contents, search_file)
    
    'VB�v���W�F�N�g�t�@�C�����Q�Ƃ��Ă���t�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
    CopyProjectFiles out_path, filelist
    
    'BAT�t�@�C�����쐬����
    CreateBatFile search_path, out_path, out_bat, filelist
    
    '�V�[�g�����w�肳��Ă���΃V�[�g��VB�v���W�F�N�g�t�@�C�����o�͂���
    'TODO

End Sub

'���W�����������؂���
Private Function Validate() As String
    If search_file = "" Or _
       search_path = "" Or _
       encode = "" Or _
       out_path = "" Then
        Validate = "�����͂̏�񂪂���܂�"
        Exit Function
    End If

    Dim ext As String: ext = Common.GetFileExtension(search_file)
    
    If ext <> "vbp" And ext <> "vbproj" Then
        Validate = "VB�v���W�F�N�g�t�@�C���������Ή��̊g���q�ł�"
        Exit Function
    End If

    If Common.IsExistsFolder(search_path) = False Then
        Validate = "�����t�H���_�����݂��܂���"
        Exit Function
    End If
    
    If out_bat <> "" Then
        ext = Common.GetFileExtension(out_bat)
        If ext <> "bat" Then
            Validate = "BAT�t�@�C���������Ή��̊g���q�ł�"
            Exit Function
        End If
    End If

    Validate = ""
End Function

'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
Private Function ParseContents(ByRef contents() As String, ByVal filename As String) As String()
    Dim ext As String: ext = Common.GetFileExtension(filename)
    
    If ext = "vbp" Then
        ParseContents = ParseVB6Project(contents)
    Else
        ParseContents = ParseVBNETProject(contents)
    End If

End Function

'vbp�t�@�C���̃p�[�X���s��
'
'vbp�t�@�C���̃p�[�X�ΏۂƓ��e�̗�͈ȉ��̒ʂ�B
'-----------------------------------------
'Module=module1; module1.bas
'Module=module2; ..\cmn\module2.bas
'Module=module3; sub\module3.bas
'Form=form1.frm
'Form=..\cmn\form2.frm
'Form=sub\form3.frm
'Class=class1; class1.cls
'Class=class2; ..\cmn\class2.cls
'Class=class3; sub\class3.cls
'ResFile32="resfile321.RES"
'ResFile32="..\cmn\resfile322.RES"
'ResFile32="sub\resfile323.RES"
'-----------------------------------------
'��L��̏ꍇ�A�ȉ��̔z�񂪕Ԃ� (base_path��C:\tmp\base�̏ꍇ)
'[0] : "C:\tmp\base\module1.bas"
'[1] : "C:\tmp\cmn\module2.bas"
'[2] : "C:\tmp\base\sub\module3.bas"
'[3] : "C:\tmp\base\form1.frm"
'[4] : "C:\tmp\cmn\form2.frm"
'[5] : "C:\tmp\base\sub\form3.frm"
'[6] : "C:\tmp\base\class1.cls"
'[7] : "C:\tmp\cmn\class2.cls"
'[8] : "C:\tmp\base\sub\class3.cls"
'[9] : "C:\tmp\base\resfile321.RES"
'[10] :"C:\tmp\cmn\resfile322.RES"
'[11] :"C:\tmp\base\sub\resfile323.RES"
'[12] :"C:\tmp\base\test.vbp"
Private Function ParseVB6Project(ByRef contents() As String) As String()
    Dim i As Integer
    Dim filelist() As String
    Dim datas() As String
    Dim key As String
    Dim value As String
    
    Dim vbp_path As String: vbp_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbp_path)

    For i = LBound(contents) To UBound(contents)
        If InStr(contents(i), "=") = 0 Then
            '"="���܂܂Ȃ��̂Ŗ���
            GoTo CONTINUE
        End If
        
        'Key/Value�ɕ�����
        datas = Split(contents(i), "=")
        
        '�L�[���擾
        key = datas(0)
        
        '�ΏۃL�[��?
        If key <> "Module" And key <> "Form" And key <> "Class" And key <> "ResFile32" Then
            '�ΏۊO�Ȃ̂Ŗ���
            GoTo CONTINUE
        End If
        
        '�l���擾
        value = Replace(datas(1), """", "")
        
        ReDim Preserve filelist(i)
        Dim path As String
        
        If InStr(value, ";") > 0 Then
            path = Trim(Split(value, ";")(1))
        Else
            path = Trim(value)
        End If
        
        '��΃p�X�ɕϊ�����
        filelist(i) = Common.GetAbsolutePathName(base_path, path)
        
CONTINUE:
    Next i
    
    '�Ō��vbp�t�@�C���R�s�[���ǉ�����
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 1)
    filelist(filelist_cnt + 1) = vbp_path
    
    ParseVB6Project = filelist
End Function

'vbproj�t�@�C���̃p�[�X���s��
'
'vbproj�t�@�C���̃p�[�X�ΏۂƓ��e�̗�͈ȉ��̒ʂ�B
'-----------------------------------------
'TBD
'-----------------------------------------
Private Function ParseVBNETProject(ByRef contents() As String) As String()
    'TBD
    ParseVBNETProject = Nothing
End Function

'VB�v���W�F�N�g�t�@�C�����Q�Ƃ��Ă���t�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
Private Sub CopyProjectFiles(ByVal dest_path As String, ByRef filelist() As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim src_base_path As String: src_base_path = Common.GetCommonString(filelist)
    Dim i As Integer
    
    For i = LBound(filelist) To UBound(filelist)
        Dim src_path As String: src_path = filelist(i)
        Dim dst_path As String: dst_path = Replace(src_path, src_base_path, dest_path & Application.PathSeparator)
        Dim path As String: path = Common.GetFolderNameFromPath(dst_path)
        
        '�t�H���_�����݂��Ȃ��ꍇ�͍쐬����
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        '�t�@�C�����R�s�[����
        fso.CopyFile src_path, dst_path
    Next i
    
    Set fso = Nothing
End Sub

'BAT�t�@�C�����쐬����
'�쐬�C���[�W (SJIS�ō쐬���邱��)
'-------------------
'@echo off
'set SRC_DIR=C:\src
'set DST_DIR=C:\dst
'
'echo SRC_DIR=%SRC_DIR%
'echo DST_DIR=%DST_DIR%
'
'REM �t�H���_�K�w�����R�s�[
'xcopy /E /I /Y /T "%SRC_DIR%" "%DST_DIR%"
'
'REM �e�t�@�C�����R�s�[
'xcopy /Y /F "%SRC_DIR%\base\module1.bas"        "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\module2.bas"         "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\module3.bas"    "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\form1.frm"          "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\form2.frm"           "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\form3.frm"      "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\class1.cls"         "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\class2.cls"          "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\class3.cls"     "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\resfile321.RES"     "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\resfile322.RES"      "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\resfile323.RES" "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\test.vbp"           "%DST_DIR%\base"
'
'pause
'-------------------
Private Sub CreateBatFile(ByVal src_path As String, ByVal dst_path As String, ByVal bat_path As String, ByRef filelist() As String)
    Dim i As Integer
    Dim contents() As String
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    Dim file, src, dst As String
    
    Const OFFSET = 13
    ReDim Preserve contents(filelist_cnt + OFFSET)
    
    '�R�}���h�쐬�J�n
    contents(0) = "@echo off"
    contents(1) = "set SRC_DIR=" & src_path
    contents(2) = "set DST_DIR=" & dst_path
    contents(3) = ""
    contents(4) = "echo SRC_DIR=%SRC_DIR%"
    contents(5) = "echo DST_DIR=%DST_DIR%"
    contents(6) = ""
    contents(7) = "REM �t�H���_�K�w�����R�s�["
    contents(8) = "xcopy /E /I /Y /T ""%SRC_DIR%"" ""%DST_DIR%"""
    contents(9) = ""
    contents(10) = "REM �e�t�@�C�����R�s�["

    For i = LBound(filelist) To UBound(filelist)
        file = filelist(i)
        src = "%SRC_DIR%" & Replace(file, src_path, "")
        dst = "%DST_DIR%" & Replace(Common.GetFolderNameFromPath(file), src_path, "")
        contents(i + OFFSET - 2) = "xcopy /Y /F " & """" & src & """" & " " & """" & dst & """"
    Next i
    
    contents(filelist_cnt + OFFSET - 1) = ""
    contents(filelist_cnt + OFFSET) = "pause"
    
    '�t�@�C���ɏo�͂���
    Common.CreateSJISTextFile contents, bat_path

End Sub


