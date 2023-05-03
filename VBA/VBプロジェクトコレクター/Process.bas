Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�O�����s�L�� (True=�O�����s)
Public IS_EXTERNAL As Boolean

'�p�����[�^
Public main_param As MainParam
Public sub_param As SubParam

Private vbprj_files() As String

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    'VB�v���W�F�N�g�t�@�C������������
    SearchVBProjFile
    
    '�R�s�[��t�H���_���폜����
    DeleteDestFolder
    
    Dim i As Integer
    Dim copy_files() As String
    
    '���C�����[�v
    For i = LBound(vbprj_files) To UBound(vbprj_files)
        Dim vbproj_path As String: vbproj_path = vbprj_files(i)
        Common.WriteLog "i=" & i & ":[" & vbproj_path & "]"
    
        'VB�v���W�F�N�g�t�@�C���̃p�[�X���s���A�R�s�[����t�@�C�����X�g���쐬����
        copy_files = CreateCopyFileList(vbproj_path)
        
        'VB�v���W�F�N�g�t�@�C�����Q�Ƃ��Ă���t�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
        Dim dst_path As String: dst_path = main_param.GetDestDirPath() & SEP & GetProjectName(vbproj_path)
        CopyProjectFiles dst_path, copy_files, vbproj_path
        
        '�R�s�[BAT�t�@�C�����쐬����
        CreateCopyBatFile vbproj_path, dst_path, copy_files
    
        'VB�v���W�F�N�g�t�@�C�����V�[�g�o�͂���
        OutputSheet vbproj_path
    Next i
    
    '�r���hBAT�t�@�C�����쐬����
    CreateBuildBatFile vbprj_files

    Common.WriteLog "Run E"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    If IS_EXTERNAL = False Then
        Set main_param = New MainParam
        Set sub_param = New SubParam
        main_param.Init
        sub_param.Init
    End If
    
    'Main Params
    main_param.Validate
    
    'Sub Params
    sub_param.Validate
    
    Common.WriteLog main_param.GetAllValue()
    
    'Main Param�ASub Param�̂ǂ���ɂ�VB�v���W�F�N�g�t�@�C�����w�肳��Ă��Ȃ��ꍇ��NG
    If main_param.GetVBPrjFileName() = "" And _
       sub_param.GetVBProjFilePathListCount() <= 0 Then
        err_msg = "VB�v���W�F�N�g�t�@�C�����w�肳��Ă��܂���B"
        Common.WriteLog "CheckAndCollectParam E3 (" & err_msg & ")"
        Err.Raise 53, , err_msg
    End If

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'VB�v���W�F�N�g�t�@�C������������
Private Sub SearchVBProjFile()
    Common.WriteLog "SearchVBProjFile S"
    
    Dim err_msg As String
    Dim path As String
    Dim i As Integer: i = 0
    
    'VB�v���W�F�N�g�t�@�C������������
    If main_param.GetVBPrjFileName() <> "" Then
        path = Common.SearchFile(main_param.GetSrcDirPath(), main_param.GetVBPrjFileName())
        ReDim Preserve vbprj_files(i)
        vbprj_files(i) = path
    End If
    
    'Sub Param�Ɏw�肳�ꂽ�p�X���}�[�W
    If sub_param.GetVBProjFilePathListCount() > 0 Then
        vbprj_files = Common.MergeArray(vbprj_files, sub_param.GetVBProjFilePathList())
    End If
    
    vbprj_files = Common.DeleteEmptyArray(vbprj_files)
    
    If Common.IsEmptyArray(vbprj_files) = True Then
        err_msg = "VB�v���W�F�N�g�t�@�C����������܂���ł���"
        Common.WriteLog "SearchVBProjFile E1 (" & err_msg & ")"
        Err.Raise 53, , err_msg
    End If
    
    Common.WriteLog "SearchVBProjFile E"
End Sub

'�R�s�[��t�H���_���폜����
Private Sub DeleteDestFolder()
    Common.WriteLog "DeleteDestFolder S"

    If Common.IsExistsFolder(main_param.GetDestDirPath()) = True Then
        Common.DeleteFolder main_param.GetDestDirPath()
    End If
    
    Common.CreateFolder main_param.GetDestDirPath()

    Common.WriteLog "DeleteDestFolder E"
End Sub

'VB�v���W�F�N�g�t�@�C���̃p�[�X���s���A�R�s�[����t�@�C�����X�g���擾����
Private Function CreateCopyFileList(ByVal vbproj_path As String) As String()
    Common.WriteLog "CreateCopyFileList S"
    
    'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
    CreateCopyFileList = ParseContents(vbproj_path)
    
    Common.WriteLog "CreateCopyFileList E"
End Function

'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
Private Function ParseContents(ByVal vbproj_path As String) As String()
    Common.WriteLog "ParseContents S"
    
    'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
    Dim contents() As String: contents = GetVBPrjContents(vbproj_path)
    
    '�����Ƀt�@�C���p�X��ǉ�����
    Dim cnt As Integer: cnt = UBound(contents)
    ReDim Preserve contents(cnt + 1)
    contents(cnt + 1) = vbproj_path
    
    If Common.GetFileExtension(vbproj_path) = "vbp" Then
        ParseContents = ParseVB6Project(contents)
    Else
        ParseContents = ParseVBNETProject(contents)
    End If

    Common.WriteLog "ParseContents E"
End Function

'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
Private Function GetVBPrjContents(ByVal vbproj_path As String) As String()
    Common.WriteLog "GetVBPrjContents S"
    
    'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
    Dim raw_contents As String: raw_contents = Common.ReadTextFileBySJIS(vbproj_path)
    
    '�t�@�C���̓��e��z��Ɋi�[����
    Dim contents() As String: contents = Split(raw_contents, vbCrLf)
    
    GetVBPrjContents = Common.DeleteEmptyArray(contents)
    
    Common.WriteLog "GetVBPrjContents E"
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
'UserControl = usercontrol1.ctl
'UserControl=..\cmn\usercontrol2.ctl
'UserControl=sub\usercontrol3.ctl
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
'[12] :"C:\tmp\base\usercontrol1.ctl
'[13] :"C:\tmp\cmn\usercontrol2.ctl
'[14] :"C:\tmp\base\sub\usercontrol3.ctl
'[15] :"C:\tmp\base\test.vbp"
Private Function ParseVB6Project(ByRef contents() As String) As String()
    Common.WriteLog "ParseVB6Project S"

    Dim i, cnt As Integer
    Dim filelist() As String
    Dim datas() As String
    Dim key As String
    Dim value As String
    
    Dim vbp_path As String: vbp_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbp_path)

    cnt = 0

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
        If key <> "Module" And key <> "Form" And key <> "Class" And key <> "ResFile32" And key <> "UserControl" Then
            '�ΏۊO�Ȃ̂Ŗ���
            GoTo CONTINUE
        End If
        
        '�l���擾
        value = Replace(datas(1), """", "")
        
        ReDim Preserve filelist(cnt)
        Dim path As String
        
        If InStr(value, ";") > 0 Then
            path = Trim(Split(value, ";")(1))
        Else
            path = Trim(value)
        End If
        
        '��΃p�X�ɕϊ�����
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
CONTINUE:
    Next i
    
    '�Ō��vbp�t�@�C�����ǉ�����
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 1)
    filelist(filelist_cnt + 1) = vbp_path
    
    ParseVB6Project = filelist
    Common.WriteLog "ParseVB6Project E"
End Function

'vbproj�t�@�C���̃p�[�X���s��
'
'vbproj�t�@�C���̃p�[�X�ΏۂƓ��e�̗�͈ȉ��̒ʂ�B
'-----------------------------------------
'<?xml version="1.0" encoding="utf-8"?>
'<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
'  <ItemGroup>
'    <Compile Include="..\cmn\cmn.vb" />
'    <Compile Include="base.vb" />
'    <Compile Include="sub\sub.vb" />
'  </ItemGroup>
'</Project>
'-----------------------------------------
'��L��̏ꍇ�A�ȉ��̔z�񂪕Ԃ� (base_path��C:\tmp\base�̏ꍇ)
'[0] : "C:\tmp\base\base.vb"
'[1] : "C:\tmp\cmn\cmn.vb"
'[2] : "C:\tmp\base\sub\sub.vb"
'[3] : "C:\tmp\base\test.vbproj"
Private Function ParseVBNETProject(ByRef contents() As String) As String()
    Common.WriteLog "ParseVBNETProject S"

    Dim i, cnt As Integer
    Dim filelist() As String
    
    Dim vbproj_path As String: vbproj_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbproj_path)

    cnt = 0

    For i = LBound(contents) To UBound(contents)
        If InStr(contents(i), "<Compile Include=") = 0 And _
           InStr(contents(i), "<EmbeddedResource Include=") = 0 And _
           InStr(contents(i), "<None Include=") = 0 And _
           InStr(contents(i), "<HintPath>") = 0 Then
            '�r���h�ɕK�v�ȃt�@�C�����܂܂Ȃ��̂Ŗ���
            GoTo CONTINUE
        End If
        
        ReDim Preserve filelist(cnt)
        
        Dim path As String
        
        If InStr(contents(i), "<Compile Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<Compile Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<EmbeddedResource Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<EmbeddedResource Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<None Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<None Include=""", ""), """ />", ""))
        Else
            path = Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", ""))
        End If
        
        path = Replace(path, """>", "")
        
        '��΃p�X�ɕϊ�����
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
CONTINUE:
    Next i
    
    '�Ō��vbproj, sln�t�@�C�����ǉ�����
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 2)
    filelist(filelist_cnt + 1) = vbproj_path
    filelist(filelist_cnt + 2) = Replace(vbproj_path, ".vbproj", ".sln")
    
    ParseVBNETProject = filelist
    Common.WriteLog "ParseVBNETProject E"
End Function

'VB�v���W�F�N�g�t�@�C�����Q�Ƃ��Ă���t�@�C���𓯂��t�H���_�\���̂܂܃R�s�[����
Private Sub CopyProjectFiles(ByVal in_dest_path As String, ByRef filelist() As String, ByVal vbprj_path As String)
    Common.WriteLog "CopyProjectFiles S"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim SEP As String: SEP = Application.PathSeparator
    Dim base_path As String: base_path = Common.GetCommonString(filelist)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim dst_file_path() As String
    Dim i As Integer
    Dim cnt As Integer: cnt = 0
    
    Common.DeleteFolder in_dest_path
    
    For i = LBound(filelist) To UBound(filelist)
        Dim src As String: src = filelist(i)
        
        If Common.GetFileExtension(src) = "sln" And _
           Common.IsExistsFile(src) = False Then
           'sln�̏ꍇ�A�R�s�[���ɑ��݂��Ȃ��ꍇ�͖�������
           Common.WriteLog "[SKIP]" & src
           GoTo CONTINUE
        End If
        
        If Common.IsExistsFile(src) = False Then
            Err.Raise 53, , "VB�v���W�F�N�g�ɋL�ڂ��ꂽ�t�@�C�������݂��܂���" & vbCrLf & _
                            "VB Project=" & vbprj_path & vbCrLf & _
                            "Not found=" & src
        End If
        
        Dim dst As String: dst = in_dest_path & SEP & dst_base_path & Replace(src, base_path, "")
        Dim path As String: path = Common.GetFolderNameFromPath(dst)
        
        '�t�H���_�����݂��Ȃ��ꍇ�͍쐬����
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        '�t�@�C�����R�s�[����
        fso.CopyFile src, dst
        
        If Common.GetFileExtension(dst) = "vbp" Then
            'VBP�t�@�C����Path32�̓R���p�C�����ɂ͕s�v�Ȃ̂ō폜���Ă���
            DeletePath32FromVBPFile dst
        End If
        
        ReDim Preserve dst_file_path(cnt)
        dst_file_path(cnt) = dst
        
        cnt = cnt + 1
        
CONTINUE:
        
    Next i
    
    '�ړ��N�_�t�H���_���ړ�����
    MoveBaseFolder in_dest_path, dst_file_path, vbprj_path
    
    Set fso = Nothing
    Common.WriteLog "CopyProjectFiles E"
End Sub

'�ړ��N�_�t�H���_���ړ�����
Private Sub MoveBaseFolder( _
    ByVal in_dest_path As String, _
    ByRef dst_file_path() As String, _
    ByVal vbprj_path As String _
)
    Common.WriteLog "MoveBaseFolder S"

    If main_param.GetMoveBaseDirName() = "" Then
        Common.WriteLog "MoveBaseFolder E1"
        Exit Sub
    End If
    
    '�ړ��N�_�t�H���_�����w�肳��Ă���ꍇ�A�R�s�[��t�H���_�p�X�ɑ��݂��邩�`�F�b�N����
    Dim base_dir As String: base_dir = ""
    Dim i As Long
    For i = LBound(dst_file_path) To UBound(dst_file_path)
        base_dir = GetFolderPathByKeyword( _
                        Common.GetFolderNameFromPath(dst_file_path(i)), _
                        main_param.GetMoveBaseDirName())
        If base_dir <> "" Then
            Exit For
        End If
    Next i
    
    '���݂��Ȃ��ꍇ�͉������Ȃ�
    If base_dir = "" Then
        Common.WriteLog "MoveBaseFolder E2"
        Exit Sub
    End If
    
    '���݂���ꍇ�̓��l�[�����Ĉړ�����
    Dim renamed_dir As String: renamed_dir = main_param.GetMoveBaseDirName() & "_" & GetProjectName(vbprj_path)
    Dim renamed_path As String: renamed_path = Common.RenameFolder(base_dir, renamed_dir)
    
    Common.MoveFolder renamed_path, main_param.GetDestDirPath() & SEP & renamed_dir
    Common.DeleteFolder in_dest_path
    
    Common.WriteLog "MoveBaseFolder E"
End Sub

'�t�H���_�p�X�Ɏw��t�H���_�������邩�`�F�b�N���A����΂��̃t�H���_�܂ł̃p�X��Ԃ�
Private Function GetFolderPathByKeyword(path As String, keyword As String) As String
    Common.WriteLog "GetFolderPathByKeyword S"
    
    Dim path_ary() As String
    Dim ret_ary() As String
    Dim i As Integer
    Dim j As Integer
    
    path_ary = Split(path, SEP)

    For i = UBound(path_ary) To 0 Step -1
        If path_ary(i) = keyword Then
        
            ReDim Preserve ret_ary(i)
            
            For j = LBound(ret_ary) To UBound(ret_ary)
                ret_ary(j) = path_ary(j)
            Next j
        
            GetFolderPathByKeyword = Join(ret_ary, SEP)
            Common.WriteLog "GetFolderPathByKeyword E1"
            Exit Function
        End If
    Next i
    
    GetFolderPathByKeyword = ""
    Common.WriteLog "GetFolderPathByKeyword E"
End Function

'VB�v���W�F�N�g����Ԃ�
Private Function GetProjectName(ByVal vbprj_file_path As String) As String
    Common.WriteLog "GetProjectName S"
    Dim vbprj_file_name As String: vbprj_file_name = Common.GetFileName(vbprj_file_path)
    Dim ext As String: ext = Common.GetFileExtension(vbprj_file_name)
    GetProjectName = Replace(vbprj_file_name, "." & ext, "")
    Common.WriteLog "GetProjectName E"
End Function

'�R�s�[BAT�t�@�C�����쐬����
'�쐬�C���[�W (SJIS�ō쐬���邱��)
'-------------------
'@echo off
'set SRC_DIR=C:\src
'set DST_DIR=C:\_tmp
'
'echo SRC_DIR=%SRC_DIR%
'echo DST_DIR=%DST_DIR%
'
'REM �e�t�@�C�����R�s�[
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\module1.bas" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\module2.bas" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\module3.bas" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\form1.frm" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\form2.frm" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\form3.frm" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\class1.cls" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\class2.cls" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\class3.cls" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\resfile321.RES" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\resfile322.RES" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\resfile323.RES" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\usercontrol1.ctl" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\usercontrol2.ctl" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\usercontrol3.ctl" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\test.vbp" "%DST_DIR%\C\src\base"
'
'
'pause
'-------------------
Private Sub CreateCopyBatFile( _
    ByVal vbproj_path As String, _
    ByVal dst_path As String, _
    ByRef copy_files() As String _
)
    Common.WriteLog "CreateCopyBatFile S"

    If main_param.IsCreateCopyBat() = False Then
        Common.WriteLog "CreateCopyBatFile E1"
        Exit Sub
    End If
    
    Dim i As Long
    Dim contents() As String
    Dim contents_cnt As Long
    Dim base_path As String: base_path = Common.GetCommonString(copy_files)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim bat_name As String: bat_name = GetProjectName(vbproj_path) & ".bat"

    Const FIRST_ROW_CNT = 7
    Const row_cnt = 3
    Const SECOND_ROW_CNT = 2
    
    ReDim Preserve contents(FIRST_ROW_CNT)
    
    '�R�}���h�쐬�J�n
    contents(0) = "@echo off"
    contents(1) = "set SRC_DIR=" & Common.RemoveTrailingBackslash(base_path)
    contents(2) = "set DST_DIR=" & dst_path
    contents(3) = ""
    contents(4) = "echo SRC_DIR=%SRC_DIR%"
    contents(5) = "echo DST_DIR=%DST_DIR%"
    contents(6) = ""
    contents(7) = "REM �e�t�@�C�����R�s�["
    
    Dim OFFSET As Long: OFFSET = UBound(contents) + 1

    For i = LBound(copy_files) To UBound(copy_files)
        contents_cnt = UBound(contents)
        ReDim Preserve contents(contents_cnt + row_cnt)
    
        Dim file As String: file = copy_files(i)
        
        Dim src As String: src = "%SRC_DIR%" & SEP & Replace(file, base_path, "")
        Dim dst_tmp As String: dst_tmp = "%DST_DIR%" & SEP & dst_base_path & Replace(file, base_path, "")
        Dim dst As String: dst = Common.GetFolderNameFromPath(dst_tmp)
        
        contents(i * row_cnt + OFFSET) = "md " & DQ & dst & DQ
        contents(i * row_cnt + OFFSET + 1) = "xcopy /Y /F " & DQ & src & DQ & " " & DQ & dst & DQ
        contents(i * row_cnt + OFFSET + 2) = ""
    Next i
    
    contents_cnt = UBound(contents)
    ReDim Preserve contents(contents_cnt + SECOND_ROW_CNT)
    contents(contents_cnt + SECOND_ROW_CNT - 1) = ""
    contents(contents_cnt + SECOND_ROW_CNT) = "pause"
    
    '�t�@�C���ɏo�͂���
    Common.CreateSJISTextFile contents, dst_path & SEP & bat_name
    
    Common.WriteLog "CreateCopyBatFile E"
End Sub

'�r���hBAT�t�@�C�����쐬����
' https://stackoverflow.com/questions/3444505/what-are-the-command-line-options-for-the-vb6-ide-compiler
' https://sh-yoshida.hatenablog.com/entry/2017/05/27/012755
Private Sub CreateBuildBatFile(ByRef vbprj_files() As String)
    Common.WriteLog "CreateBuildBatFile S"

    If main_param.IsCreateBuildBat() = False Then
        Common.WriteLog "CreateBuildBatFile E1"
        Exit Sub
    End If
    
    Const VB6EXE = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe"
    Const MSBLDEXE = "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
    Const BUILDLOG = "build.log"
    
    Dim i As Long
    Dim contents() As String
    Dim contents_cnt As Long
       
    Const FIRST_ROW_CNT = 11
    Const row_cnt = 5
    Const SECOND_ROW_CNT = 2
    
    ReDim Preserve contents(FIRST_ROW_CNT)
    
    '�R�}���h�쐬�J�n
    contents(0) = "@echo off"
    contents(1) = "set VB6EXE=" & VB6EXE
    contents(2) = "set MSBLDEXE=" & MSBLDEXE
    contents(3) = "set BUILDLOG=" & BUILDLOG
    contents(4) = ""
    contents(5) = "echo VB6EXE=%VB6EXE%"
    contents(6) = "echo MSBLDEXE=%MSBLDEXE%"
    contents(7) = "echo BUILDLOG=%BUILDLOG%"
    contents(8) = ""
    contents(9) = "REM �e�v���W�F�N�g���r���h"
    contents(10) = "echo Start Build > %BUILDLOG%"
    contents(11) = ""
    
    Dim OFFSET As Long: OFFSET = UBound(contents) + 1

    'VB6.exe�̑��݃`�F�b�N
    'MSBuild.exe�̑��݃`�F�b�N

    '���ʃ��O�t�@�C���̑��݃`�F�b�N
    ' �����݂���ꍇ�͍폜
    
    'VB�v���W�F�N�g���[�v
    For i = LBound(vbprj_files) To UBound(vbprj_files)
        Dim path As String: path = vbprj_files(i)
        Dim ext As String: ext = Common.GetFileExtension(path)
        Dim renamed_dir As String: renamed_dir = main_param.GetMoveBaseDirName() & "_" & GetProjectName(path)
        Dim dst_path As String: dst_path = Replace(Common.GetStringByKeyword(path, main_param.GetMoveBaseDirName()), main_param.GetMoveBaseDirName() & SEP, renamed_dir & SEP)
        
        'D:\src_testVB6\base\testVB6.vbp
        Dim target_path As String: target_path = "D:\" & dst_path
        
        contents_cnt = UBound(contents)
        ReDim Preserve contents(contents_cnt + row_cnt)
        
        If ext = "vbp" Then
            
            'VB6�ŃR���p�C��
            contents(i * row_cnt + OFFSET + 0) = "IF EXIST " & DQ & "%VB6EXE%" & DQ & " ("
            contents(i * row_cnt + OFFSET + 1) = "  echo VB6 Build [" & target_path & "] >> %BUILDLOG%"
            contents(i * row_cnt + OFFSET + 2) = "  " & DQ & "%VB6EXE%" & DQ & " /m " & DQ & target_path & DQ & " /out " & "%BUILDLOG%"
            contents(i * row_cnt + OFFSET + 3) = ")"
            contents(i * row_cnt + OFFSET + 4) = ""
        
        ElseIf ext = "vbproj" Then
            
            'MSBuild�Ńr���h
            contents(i * row_cnt + OFFSET + 0) = "IF EXIST " & DQ & "%MSBLDEXE%" & DQ & " ("
            contents(i * row_cnt + OFFSET + 1) = "  echo VB.NET Build [" & target_path & "] >> %BUILDLOG%"
            contents(i * row_cnt + OFFSET + 2) = "  " & DQ & "%MSBLDEXE%" & DQ & " " & DQ & Replace(target_path, "D:\", "C:\") & DQ & " /t:clean;rebuild /p:Configuration=Release /fl"
            contents(i * row_cnt + OFFSET + 3) = ")"
            contents(i * row_cnt + OFFSET + 4) = ""
        
        End If
        
    Next i

    contents_cnt = UBound(contents)
    ReDim Preserve contents(contents_cnt + SECOND_ROW_CNT)
    contents(contents_cnt + SECOND_ROW_CNT - 1) = ""
    contents(contents_cnt + SECOND_ROW_CNT) = "pause"
    
    '�t�@�C���ɏo�͂���
    Common.CreateSJISTextFile contents, main_param.GetDestDirPath() & SEP & "Build_" & Common.GetNowTimeString() & ".bat"
    
    Common.WriteLog "CreateBuildBatFile E"
End Sub

'VBP�t�@�C����Path32�̓R���p�C�����ɂ͕s�v�Ȃ̂ō폜���Ă���
Private Sub DeletePath32FromVBPFile(ByVal path As String)
    Common.WriteLog "DeletePath32FromVBPFile S"

    If main_param.IsDeletePath32() = False Then
        Common.WriteLog "DeletePath32FromVBPFile E1"
        Exit Sub
    End If
    
    Common.RemoveLinesWithKeyword path, "Path32="

    Common.WriteLog "DeletePath32FromVBPFile E"
End Sub

'VB�v���W�F�N�g�t�@�C�����V�[�g�o�͂���
Private Sub OutputSheet(ByVal vbproj_path As String)
    If IS_EXTERNAL = True Then
        Exit Sub
    End If

    Common.WriteLog "OutputSheet S"

    If main_param.IsOutSheet() = False Then
        Common.WriteLog "OutputSheet E1"
        Exit Sub
    End If
    
    Dim sheet_name As String: sheet_name = GetProjectName(vbproj_path)
    
    'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
    Dim contents() As String: contents = GetVBPrjContents(vbproj_path)
    
    Dim prj_path As String: prj_path = contents(UBound(contents))
    
    Dim before_sheet_name As String: before_sheet_name = ActiveSheet.name
    
    Common.AddSheet sheet_name
    
    '�t�@�C���̓��e���w�肳�ꂽ�V�[�g�ɏo�͂���
    Common.OutputTextFileToSheet vbproj_path, sheet_name
    
    ThisWorkbook.Sheets(before_sheet_name).Select
    
    Common.WriteLog "OutputSheet E"
End Sub
