Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Public main_param As MainParam
Public sub_param As SubParam

Private target_files() As String
Private success_cnt As Long

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Erase target_files
    success_cnt = 0

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    '�R�s�[��t�H���_���폜����
    DeleteDestFolder
    
    Dim i As Long
    
    '���C�����[�v
    For i = LBound(target_files) To UBound(target_files)
        '�R�s�[
        CopyFiles target_files(i)
    Next i

    Common.WriteLog "Run E"
End Sub

Public Function GetResult() As String
    GetResult = success_cnt & "/" & sub_param.GetFilePathListCount()
End Function

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Set main_param = New MainParam
    Set sub_param = New SubParam
    main_param.Init
    sub_param.Init
    
    'Main Params
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()
    
    'Sub Params
    sub_param.Validate
    
    target_files = sub_param.GetFilePathList()

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'�R�s�[��t�H���_���폜����
Private Sub DeleteDestFolder()
    Common.WriteLog "DeleteDestFolder S"

    Dim dst_path As String: dst_path = main_param.GetDestDirPath()

    If Common.IsExistsFolder(dst_path) = True Then
        If Common.IsEmptyFolder(dst_path) = False Then
            If Common.ShowYesNoMessageBox( _
                "�R�s�[��t�H���_����ł͂���܂���B" & vbCrLf & _
                "�����𑱂��܂����H" & vbCrLf & _
                "�i������ƃt�H���_�͍폜����܂�!�j" _
            ) = False Then
                Err.Raise 53, , "�R�s�[��t�H���_����ł͖����̂ŏ������L�����Z�����܂����B(" & dst_path & ")"
            End If
        End If
    End If

    If Common.IsExistsFolder(dst_path) = True Then
        Common.DeleteFolder dst_path
    End If
    
    Common.CreateFolder dst_path

    Common.WriteLog "DeleteDestFolder E"
End Sub

'�t�@�C�����R�s�[����
Private Sub CopyFiles(ByVal target_path As String)
    Common.WriteLog "CopyFiles S"
    Common.WriteLog "[CopyFiles] SrcFilePath=" & target_path
        
    Dim err_msg As String
    Dim dest_dir_path As String: dest_dir_path = main_param.GetDestDirPath()
    Dim is_copy_dir As Boolean: is_copy_dir = main_param.IsCopyDir()
    Dim is_continue As Boolean: is_continue = main_param.IsContinue()
    Dim is_overwrite As Boolean: is_overwrite = main_param.IsOverWrite
    Dim src_dir_path As String
    Dim dst_file_path As String
    
    If Common.IsExistsFile(target_path) = False Then
        err_msg = "�R�s�[���t�@�C�������݂��܂���(" & target_path & ")"
        Common.WriteLog "[CopyFiles] ��Error! " & err_msg
        
        If is_continue = False Then
            '�R�s�[�������݂��Ȃ��ꍇ�͖������đ��s���Ȃ�
            Err.Raise 53, , err_msg
        End If
        
        Common.WriteLog "CopyFiles E-1"
        Exit Sub
    End If
    
    If is_copy_dir = False Then
        '�t�H���_�̓R�s�[���Ȃ��ꍇ
        
        '�R�s�[��t�@�C���p�X���쐬
        dst_file_path = dest_dir_path & SEP & Common.GetFileName(target_path)
        
        Common.WriteLog "[CopyFiles] DestFilePath=" & target_path
        
        If Common.IsExistsFile(dst_file_path) = True Then
            '���łɓ����t�@�C��������ꍇ
            
            If is_overwrite = False Then
                '�R�s�[��ɓ����t�@�C��������ꍇ�͏㏑�����Ȃ��ꍇ�́A���j�[�N�ȃt�@�C�����ɕύX����
                Common.CopyUniqueFile target_path, dest_dir_path
                
                success_cnt = success_cnt + 1
                Common.WriteLog "CopyFiles E-2"
                Exit Sub
            End If
        End If
        
        '�㏑���R�s�[
        Common.CopyFile target_path, dst_file_path
        
        success_cnt = success_cnt + 1
        Common.WriteLog "CopyFiles E-3"
        Exit Sub
    End If
    
    '�t�H���_���R�s�[����ꍇ
    
    '�R�s�[��t�H���_�p�X���R�s�[������擾
    src_dir_path = Replace(Common.GetFolderPath(target_path), ":", "")
    
    '�R�s�[��t�@�C���p�X���쐬
    dst_file_path = dest_dir_path & SEP & Replace(target_path, ":", "")
    
    Common.WriteLog "[CopyFiles] DestFilePath=" & dst_file_path
    
    '�R�s�[
    Common.CopyFile target_path, dst_file_path, True
    
    success_cnt = success_cnt + 1
    Common.WriteLog "CopyFiles E"
End Sub
