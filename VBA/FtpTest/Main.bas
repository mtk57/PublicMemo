Attribute VB_Name = "Main"
#If VBA7 Then
  Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As LongPtr
  Declare PtrSafe Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As LongPtr) As LongPtr
#Else
  'Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
  'Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
#End If

Public Sub TEST_Click()

On Error GoTo ErrHandler

    Call ExecuteFTPScript

    Exit Sub

ErrHandler:
    MsgBox "ERROR!  " & Err.Description

End Sub

Public Sub TES2_Click()

On Error GoTo ErrHandler

    Call ExecuteFTPScript_2

    Exit Sub

ErrHandler:
    MsgBox "ERROR!  " & Err.Description

End Sub

Private Sub ExecuteFTPScript()

    ' FTP�T�[�o�[���
    Const FTP_SERVER As String = "localhost"
    Const FTP_USER As String = "ftpuser"
    Const FTP_PASSWORD As String = "ftpuser"
    Const REMOTE_FILE As String = "/test.txt"   ' �_�E�����[�h���郊���[�g�t�@�C���p�X
    Const LOCAL_FILE As String = "C:\_git\PublicMemo\VBA\FtpTest\local_test.txt"
    
    Dim scriptPath As String
    Dim scriptContent As String
    Dim cmd As String
    
    Dim ftpPath As String
    Dim result As LongPtr ' LongPtr�^�ɕύX
    
    ' ftp.exe�̃t���p�X���擾
    ftpPath = Environ("SystemRoot") & "\System32\ftp.exe"
    
    ' �X�N���v�g�t�@�C���p�X���쐬 (�ꎞ�t�@�C���Ƃ��č쐬)
    scriptPath = Environ("TEMP") & "\ftp_script.txt"
    
    ' �X�N���v�g���e���쐬
    scriptContent = "open " & FTP_SERVER & vbCrLf & _
                    FTP_USER & vbCrLf & _
                    FTP_PASSWORD & vbCrLf & _
                    "ascii" & vbCrLf & _
                    "get " & REMOTE_FILE & " " & LOCAL_FILE & vbCrLf & _
                    "bye"
    
    ' �X�N���v�g�t�@�C����ۑ�
    Open scriptPath For Output As #1
      Print #1, scriptContent
    Close #1
    
    ' ftp.exe�����s (�p�X�ɃX�y�[�X���܂܂��ꍇ���l������""�ň͂�)
    cmd = """" & ftpPath & """ -s:" & scriptPath
    
    ' �J�����g�f�B���N�g�����ꎞ�t�H���_�ɐݒ� (�d�v)
    ChDir Environ("TEMP")
    
    ' WinExec�Ŏ��s
    result = WinExec(cmd, 0) ' 0 �� vbHide �Ɠ���
    
    Select Case result
      Case 0: MsgBox "FTP�R�}���h�̎��s�Ɏ��s���܂����Bftp.exe��������Ȃ��\��������܂��B", vbCritical
      Case 1 To 31: MsgBox "FTP�R�}���h���s���ɃG���[���������܂����B�G���[�R�[�h�F" & result, vbCritical
      Case Else: MsgBox "FTP�������������܂����B", vbInformation
    End Select
    
    ' (�I�v�V����) �X�N���v�g�t�@�C�����폜 (�K�v�ɉ����ăR�����g�A�E�g)
    Kill scriptPath

End Sub

'PowerShell & ���ϐ���
Private Sub ExecuteFTPScript_2()

    ' �ꎞ�I�Ȋ��ϐ���
    Const TEMP_USER_VAR As String = "TEMP_FTP_USER"
    Const TEMP_PASSWORD_VAR As String = "TEMP_FTP_PASSWORD"
    Const TEMP_SERVER_VAR As String = "TEMP_FTP_SERVER"
    Const TEMP_REMOTE_FILE_VAR As String = "TEMP_FTP_REMOTE_FILE"
    Const TEMP_LOCAL_FILE_VAR As String = "TEMP_FTP_LOCAL_FILE"

    ' FTP���
    Dim ftpUser As String
    Dim ftpPassword As String
    Dim ftpServer As String
    Dim remoteFile As String
    Dim localFile As String

    ftpUser = "ftpuser"
    ftpPassword = "ftpuser"
    ftpServer = "localhost"
    remoteFile = "/test.txt"
    localFile = "C:\_git\PublicMemo\VBA\FtpTest\local_test2.txt"
    
    ' ���ϐ���ݒ�
    Dim result As LongPtr
    result = SetEnvironmentVariable(TEMP_USER_VAR, ftpUser)
    result = SetEnvironmentVariable(TEMP_PASSWORD_VAR, ftpPassword)
    result = SetEnvironmentVariable(TEMP_SERVER_VAR, ftpServer)
    result = SetEnvironmentVariable(TEMP_REMOTE_FILE_VAR, remoteFile)
    result = SetEnvironmentVariable(TEMP_LOCAL_FILE_VAR, localFile)

    'Debug.Print "User: " & ftpUser
    'Debug.Print "Password: " & ftpPassword
    'Debug.Print "Server: " & ftpServer
    'Debug.Print "Remote File: " & remoteFile
    'Debug.Print "Local File: " & localFile

    ' PowerShell�X�N���v�g�̓��e��VBA�Ő���
    Dim scriptContent As String
    scriptContent = _
        "$username = $env:" & TEMP_USER_VAR & vbCrLf & _
        "$password = $env:" & TEMP_PASSWORD_VAR & vbCrLf & _
        "$server = $env:" & TEMP_SERVER_VAR & vbCrLf & _
        "$remoteFile = $env:" & TEMP_REMOTE_FILE_VAR & vbCrLf & _
        "$localFile = $env:" & TEMP_LOCAL_FILE_VAR & vbCrLf & _
        "$uri = ""ftp://${username}:${password}@${server}${remoteFile}""" & vbCrLf & _
        "try {" & vbCrLf & _
        "  Invoke-WebRequest -Uri $uri -OutFile $localFile -UseBasicParsing" & vbCrLf & _
        "  Write-Host ""FTP�t�@�C���̃_�E�����[�h���������܂����B""" & vbCrLf & _
        "} catch {" & vbCrLf & _
        "  Write-Error ""FTP�t�@�C���̃_�E�����[�h�Ɏ��s���܂���: $($_.Exception.Message)""" & vbCrLf & _
        "  exit 1" & vbCrLf & _
        "}"


    ' PowerShell�X�N���v�g���ꎞ�t�@�C���ɕۑ�
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\ftp_script.ps1"
       
    Open scriptPath For Output As #1
      Print #1, scriptContent
    Close #1
    
    Dim cmd As String
    cmd = "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    result = objShell.Run(cmd, 0, True) ' 0: ��\��, 1: �ʏ�\��, True: �����܂őҋ@

    If result = 0 Then
        MsgBox "PowerShell�X�N���v�g������Ɋ������܂����B", vbInformation
    Else
        MsgBox "PowerShell�X�N���v�g�̎��s�Ɏ��s���܂����B�G���[�R�[�h: " & result, vbCritical
    End If

    Set objShell = Nothing
    
    ' ���ϐ����폜
    SetEnvironmentVariable TEMP_USER_VAR, vbNullString
    SetEnvironmentVariable TEMP_PASSWORD_VAR, vbNullString
    SetEnvironmentVariable TEMP_SERVER_VAR, vbNullString
    SetEnvironmentVariable TEMP_REMOTE_FILE_VAR, vbNullString
    SetEnvironmentVariable TEMP_LOCAL_FILE_VAR, vbNullString
    
    ' �X�N���v�g�t�@�C�����폜
    Kill scriptPath

End Sub



