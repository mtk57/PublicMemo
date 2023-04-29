Attribute VB_Name = "Process_002"
Option Explicit

Private prms As ParamContainer
Private SEP As String
Private DQ As String

Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Dim msg As String: msg = ""

    Set prms = New ParamContainer
    
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    Dim i As Integer
    Dim target As ParamTarget
    Dim targetlist() As ParamTarget
    targetlist = prms.GetTargetList()
        
    DoClone
    
    For i = LBound(targetlist) To UBound(targetlist)
    
        Set target = targetlist(i)
    
        SwitchDevelopBranch
        
        DoPull
        
        CreateFeatureBranch target
        
        DoCopy target
        
        DoAdd
        
        DoCommit target
        
        DoTag target
        
        DoPush target
    
    Next i
        
    Common.WriteLog "Run E"
End Sub

Private Sub DoClone()
    Common.WriteLog "DoClone S"
    
    If prms.GetUrl() = "" Then
        Common.WriteLog "DoClone E1"
        Exit Sub
    End If
    
    If Common.IsExistsFolder(prms.GetGitDirPath()) = True Then
        Common.WriteLog "DoClone E2"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    '�N���[��
    cmd = "git clone " & prms.GetUrl() & " " & prms.GetGitDirPath()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoClone E"
End Sub

Private Sub SwitchDevelopBranch()
    Common.WriteLog "SwitchDevelopBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'develop�u�����`�����݂��Ȃ��ꍇ�̓G���[�Ƃ���
    If IsExistBranch(prms.GetBaseBranch()) = False Then
        Err.Raise 53, , "�u�����`�����݂��܂���B(" & prms.GetBaseBranch() & ")"
    End If
    
    '�J�����g�u�����`���m�F����
    cmd = "git branch --show-current"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If git_result(0) = prms.GetBaseBranch() Then
        Common.WriteLog "CheckoutBranch E1"
        Exit Sub
    End If
    
    '���[�J���u�����`�����邩�m�F����
    Dim is_exist_local As Boolean: is_exist_local = False
    
    cmd = "git branch --list " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If Common.IsEmptyArray(git_result) = False Then
        is_exist_local = True
    End If
    
    If is_exist_local = True Then
        '���[�J���u�����`������̂�switch�Ő؂�ւ�
        cmd = "git switch " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    Else
        '���[�J���u�����`���Ȃ��̂ō쐬���Đ؂�ւ�
        cmd = "git checkout -b " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If
    
    Common.WriteLog "SwitchDevelopBranch E"
End Sub

Private Sub DoPull()
    Common.WriteLog "DoPull S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�v��
    cmd = "git pull origin " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPull E"
End Sub

Private Sub CreateFeatureBranch(ByRef target As ParamTarget)
    Common.WriteLog "CreateFeatureBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    If IsExistBranch(target.GetBranch()) = True Then
        If prms.IsDeleteExistBranch() = False Then
            'feature�u�����`�����ɑ��݂���ꍇ�̓G���[�Ƃ���
            Err.Raise 53, , "�u�����`�����ɑ��݂��܂��B(" & target.GetBranch() & ")"
        Else
            'feature�u�����`���폜
            cmd = "git branch --delete " & target.GetBranch()
            git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
        End If
    End If
    
    'feature�u�����`���쐬���Đ؂�ւ�
    cmd = "git checkout -b " & target.GetBranch() & " " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "CreateFeatureBranch E"
End Sub

'VB�v���W�F�N�g�����W�����t�H���_��������A�R�s�[����VB�v���W�F�N�g�t�@�C���܂ł̃t�H���_�p�X��
'�N�_�t�H���_����܂߂Ċ��S��v����t�H���_�������āAGit�t�H���_�ɃR�s�[����
Private Sub DoCopy(ByRef target As ParamTarget)
    Common.WriteLog "DoCopy S"
    
    '�R�s�[����VB�v���W�F�N�g�t�@�C���܂ł̃t�H���_�p�X���N�_�t�H���_����t�@�C�����܂ł��擾����
    Dim path As String: path = Common.GetStringByKeyword(target.GetVBPrjFilePath(), SEP & prms.GetBaseFolder() & SEP)
    
    '�N�_�t�H���_�����l�[��
    Dim prj_name As String: prj_name = GetProjectName(path)
    path = Replace(path, SEP & prms.GetBaseFolder() & SEP, prms.GetBaseFolder() & "_" & prj_name & SEP)
    
    'VB�v���W�F�N�g�����W�����t�H���_��������A��v����t�@�C�������邩�`�F�b�N����
    Dim ext As String: ext = Common.GetFileExtension(path)
    Dim file_list() As String: file_list = Common.CreateFileList(prms.GetDestDirPath(), "*." & ext, True)
    
    Dim i As Long
    Dim check_path As String
    Dim is_match As Boolean: is_match = False
    
    For i = LBound(file_list) To UBound(file_list)
        check_path = file_list(i)
        If InStr(check_path, path) > 0 Then
            is_match = True
            Exit For
        End If
    Next i
    
    If is_match = False Then
        Err.Raise 53, , "VB�v���W�F�N�g�t�@�C����������܂���B(" & check_path & ")"
    End If
    
    '�N�_�t�H���_�����l�[�����āAGit�t�H���_�ɃR�s�[
    Dim src_path As String: src_path = prms.GetDestDirPath() & SEP & prms.GetBaseFolder() & "_" & prj_name
    
    Dim dst_path As String: dst_path = prms.GetGitDirPath() & SEP
    If ext = "vbp" Then
        dst_path = dst_path & "SRC_020"
    Else
        dst_path = dst_path & "SRC_030"
    End If
    dst_path = dst_path & SEP & prms.GetBaseFolder()
    
    Common.CopyFolder src_path, dst_path
    
    Common.WriteLog "DoCopy E"
End Sub

Private Sub DoAdd()
    Common.WriteLog "DoAdd S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�C���f�b�N�X�ɒǉ�����
    cmd = "git add ."
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoAdd E"
End Sub

Private Sub DoCommit(ByRef target As ParamTarget)
    Common.WriteLog "DoCommit S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�R�~�b�g����
    cmd = "git commit -m " & DQ & target.GetCommit() & DQ
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoCommit E"
End Sub

Private Sub DoTag(ByRef target As ParamTarget)
    Common.WriteLog "DoTag S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�^�O��t����
    cmd = "git tag " & target.GetTag() & " HEAD"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoTag E"
End Sub

Private Sub DoPush(ByRef target As ParamTarget)
    Common.WriteLog "DoPush S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '�^�O��t����
    cmd = "git push --tags --set-upstream origin " & target.GetBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPush E"
End Sub

'�u�����`�̑��݃`�F�b�N
Private Function IsExistBranch(ByVal branch As String) As Boolean
    Common.WriteLog "DoIsExistBranchPush S"
    
    Dim cmd As String
    Dim git_result() As String
    Dim i As Long
    Dim is_find As Boolean: is_find = False
    
    '�S�Ẵu�����`�̃��X�g���擾����
    cmd = "git branch --all"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If Common.IsEmptyArray(git_result) = True Then
        Err.Raise 53, , "�u�����`��1�����݂��܂���B"
    End If
    
    '�u�����`�����݂��邩�`�F�b�N����
    For i = LBound(git_result) To UBound(git_result)
        If InStr(git_result(i), branch) > 0 Then
            is_find = True
            Exit For
        End If
    Next i
    
    IsExistBranch = is_find
    
    Common.WriteLog "IsExistBranch E"
End Function

'VB�v���W�F�N�g����Ԃ�
Private Function GetProjectName(ByVal vbprj_file_path As String) As String
    Common.WriteLog "GetProjectName S"
    Dim vbprj_file_name As String: vbprj_file_name = Common.GetFileName(vbprj_file_path)
    Dim ext As String: ext = Common.GetFileExtension(vbprj_file_name)
    GetProjectName = Replace(vbprj_file_name, "." & ext, "")
    Common.WriteLog "GetProjectName E"
End Function
