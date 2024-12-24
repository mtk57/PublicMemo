Attribute VB_Name = "Main"
#If VBA7 Then ' VBA7�ȍ~ (Office 2010�ȍ~) �̏ꍇ

  ' 64bit�݊���WinExec�錾
  Declare PtrSafe Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As LongPtr) As LongPtr

#Else ' VBA6�ȑO (Office 2007�ȑO) �̏ꍇ (�݊����̂��߂Ɏc��)

  'Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

#End If

Public Sub TEST_Click()

On Error GoTo ErrHandler

    Call ExecuteFTPScript

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



