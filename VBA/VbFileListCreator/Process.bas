Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

Const EXT_VBP = "*.vbp"
Const EXT_VBPROJ = "*.vbproj"

'�p�����[�^
Public main_param As MainParam

Private vbprj_files() As String
Private sheet_name As String

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Erase vbprj_files

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    'VB�v���W�F�N�g�t�@�C������������
    SearchVBProjFile
    
    '�V�K�V�[�g���쐬����
    CreateSheet
    
    Dim i As Long
    Dim j As Long
    Dim cur_row As Long: cur_row = 1
    
    '���C�����[�v
    For i = LBound(vbprj_files) To UBound(vbprj_files)
    
        Dim vbproj_path As String: vbproj_path = vbprj_files(i)
        Common.WriteLog "i=" & i & ":[" & vbproj_path & "]"
    
        'VB�v���W�F�N�g�t�@�C���̃p�[�X���s���A�o�͂���t�@�C�����X�g���쐬����
        Dim ref_files As RefFiles
        Set ref_files = CreateVbRefFileList(vbproj_path)
        
        '�V�[�g�ɏo�͂���
        Dim vbp_path As String: vbp_path = ref_files.GetSrcDirPath()
        
        For j = 0 To ref_files.GetAppendRefFileCount
            Common.UpdateSheet ActiveWorkbook, sheet_name, cur_row, 1, vbp_path
            Common.UpdateSheet ActiveWorkbook, sheet_name, cur_row, 2, ref_files.GetRefFile(j)
            cur_row = cur_row + 1
        Next j
    Next i

    Common.WriteLog "Run E"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    Set main_param = New MainParam
    main_param.Init

    'Main Params
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

'VB�v���W�F�N�g�t�@�C������������
Private Sub SearchVBProjFile()
    Common.WriteLog "SearchVBProjFile S"
    
    Dim err_msg As String
    Dim path As String
    Dim i As Integer: i = 0
    
    Dim target_ext As String
    target_ext = "*." & main_param.GetTargetType()
    
    'VB�v���W�F�N�g�t�@�C������������
    vbprj_files = Common.CreateFileList( _
        main_param.GetSrcDirPath(), _
        target_ext, _
        True)

    vbprj_files = Common.DeleteEmptyArray(vbprj_files)
    
    If Common.IsEmptyArray(vbprj_files) = True Then
        err_msg = "VB�v���W�F�N�g�t�@�C����������܂���ł���"
        Common.WriteLog "SearchVBProjFile E1 (" & err_msg & ")"
        Err.Raise 53, , err_msg
    End If
    
    Common.WriteLog "SearchVBProjFile E"
End Sub


'VB�v���W�F�N�g�t�@�C���̃p�[�X���s���A�t�@�C�����X�g���擾����
Private Function CreateVbRefFileList(ByVal vbproj_path As String) As RefFiles
    Common.WriteLog "CreateVbRefFileList S"
    
    'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
    Set CreateVbRefFileList = ParseContents(vbproj_path)
    
    Common.WriteLog "CreateVbRefFileList E"
End Function

'VB�v���W�F�N�g�t�@�C���̃p�[�X���s��
Private Function ParseContents(ByVal vbproj_path As String) As RefFiles
    Common.WriteLog "ParseContents S"
    
    'VB�v���W�F�N�g�t�@�C���̓��e��ǂݍ���
    Dim contents() As String: contents = GetVBPrjContents(vbproj_path)
    
    '�����Ƀt�@�C���p�X��ǉ�����
    Dim cnt As Integer: cnt = UBound(contents)
    ReDim Preserve contents(cnt + 1)
    contents(cnt + 1) = vbproj_path
    
    If Common.GetFileExtension(vbproj_path) = "vbp" Then
        Set ParseContents = ParseVB6Project(contents)
    Else
        Set ParseContents = ParseVBNETProject(contents)
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
Private Function ParseVB6Project(ByRef contents() As String) As RefFiles
    Common.WriteLog "ParseVB6Project S"

    Dim ref_files As RefFiles
    Set ref_files = New RefFiles

    Dim i, cnt As Integer
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
        
        Dim path As String
        
        If InStr(value, ";") > 0 Then
            path = Trim(Split(value, ";")(1))
        Else
            path = Trim(value)
        End If
        
        '��΃p�X�ɕϊ�����
        Dim abs_path As String: abs_path = Common.GetAbsolutePathName(base_path, path)
        ref_files.AppendRefFilePath (abs_path)
        cnt = cnt + 1
        
CONTINUE:
    Next i
    
    '�Ō��vbp�t�@�C�����ǉ�����
    ref_files.SetSrcDirPath (vbp_path)
    
    Set ParseVB6Project = ref_files
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
Private Function ParseVBNETProject(ByRef contents() As String) As RefFiles
    Common.WriteLog "ParseVBNETProject S"

    Dim ref_files As RefFiles
    Set ref_files = New RefFiles

    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    Dim filelist() As String
    
    Dim vbproj_path As String: vbproj_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbproj_path)
    
    '���O�t�@�C�����X�g���쐬
    Dim ignore_files() As String
    ignore_files = Split(main_param.GetIgnoreFiles(), ",")

    cnt = 0

    For i = LBound(contents) To UBound(contents)
        'Common.WriteLog "contents(" & i & ")=" & contents(i)
    
        If InStr(contents(i), "<Compile Include=") = 0 And _
           InStr(contents(i), "<EmbeddedResource Include=") = 0 And _
           InStr(contents(i), "<None Include=") = 0 And _
           InStr(contents(i), "<HintPath>") = 0 And _
           InStr(contents(i), "<ApplicationIcon>") = 0 Then
            '�r���h�ɕK�v�ȃt�@�C�����܂܂Ȃ��̂Ŗ���
            'Common.WriteLog "Skip contents(" & i & ")=" & contents(i)
            GoTo CONTINUE
        End If
        
        If Common.IsEmptyArray(ignore_files) = False Then
            For j = LBound(ignore_files) To UBound(ignore_files)
                If InStr(contents(i), ignore_files(j)) > 0 Then
                    '���O�t�@�C�����܂ނ̂Ŗ���
                    GoTo CONTINUE
                End If
            Next j
        End If
        
        If Common.StartsWith(Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", "")), "packages") Then
            'packages�͖�������
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
        ElseIf InStr(contents(i), "<HintPath>") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", ""))
        ElseIf InStr(contents(i), "<ApplicationIcon>") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<ApplicationIcon>", ""), "</ApplicationIcon>", ""))
        End If
        
        If path = "" Then
            GoTo CONTINUE
        End If
        
        path = Replace(path, """>", "")
        
        '��΃p�X�ɕϊ�����
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        ref_files.AppendRefFilePath (filelist(cnt))
        cnt = cnt + 1
        
        'ActiveReport ���ꏈ��
        If InStr(contents(i), "<Compile Include=""reports\") > 0 Then
            'rpx�̑��݃`�F�b�N���s���A����Βǉ�����
            Dim rpx_path As String: rpx_path = Replace(path, ".vb", ".rpx")
            Dim rpx_find_path As String: rpx_find_path = base_path & SEP & rpx_path
            If Common.IsExistsFile(rpx_find_path) = True Then
                Common.WriteLog "rpx found.(" & rpx_find_path & ")"
                
                ReDim Preserve filelist(cnt)
                filelist(cnt) = Common.GetAbsolutePathName(base_path, rpx_path)
                cnt = cnt + 1
            Else
                Common.WriteLog "rpx not found.(" & rpx_find_path & ")"
            End If
        End If
        
CONTINUE:
    Next i
    
    '�Ō��vbproj, sln�t�@�C�����ǉ�����
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 2)
    filelist(filelist_cnt + 1) = vbproj_path
    filelist(filelist_cnt + 2) = Replace(vbproj_path, ".vbproj", ".sln")
    
    ref_files.SetSrcDirPath (vbproj_path)
    
    'ParseVBNETProject = filelist
    Set ParseVBNETProject = ref_files
    
    Common.WriteLog "ParseVBNETProject E"
End Function

Private Sub CreateSheet()
    Common.WriteLog "CreateSheet S"
    sheet_name = Common.GetNowTimeString_OLD()
    Common.AddSheet ThisWorkbook, sheet_name
    Common.WriteLog "CreateSheet E"
End Sub
