Attribute VB_Name = "Process_002"
Option Explicit

Private prm As Param

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prm = New Param
    
    prm.Init
    prm.Validate
    
    Common.WriteLog prm.GetAllValue()
    
    Clone
    
    Common.WriteLog "Run E"
End Sub

Private Sub Clone()
    Common.WriteLog "Clone S"
    
    RunExe "git log --oneline"
    'RunExe "git log --oneline > C:\_tmp\aaa.txt"
    Dim ret
    
    'ChDir prm.GetGitDirPath()
    'ret = Shell("""git log --oneline > C:\_tmp\aaa.txt""", 1)
    
    'GetGitLog
    
    Common.WriteLog "Clone E"
End Sub

Private Sub RunExe(ByVal command As String)
    Common.WriteLog "RunExe S"

    Dim i As Integer
    Dim ret As Long
    Dim exe_param As String
        
    ChDir prm.GetGitDirPath()
      
    Common.WriteLog command
    
    'ret = Common.RunProcessWait(command)
    
    ret = RunProcessWait(command)
    
    If ret <> 0 Then
        Common.WriteLog "exe ret=" & ret
        Err.Raise 53, , "Exe�̎��s�Ɏ��s���܂���(ret=" & ret & ")"
    End If

    Common.WriteLog "RunExe E"
End Sub

'-------------------------------------------------------------
'�O���A�v���P�[�V���������s���A�I������܂őҋ@����
' exe_path : IN : �O���A�v���P�[�V����(exe)�̐�΃p�X
'                 exe�ɓn���p�����[�^������ꍇ���ꏏ�ɏ�������
' Ret : �v���Z�X�̖߂�l
'-------------------------------------------------------------
Public Function RunProcessWait(ByVal exe_path As String) As Long

    'testsub5
    
    'RunProcessWait = 0
    'Exit Function


    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    
    Const NOT_DISP = 0
    Const DISP = 1
    Const WAIT = True
    Const NO_WAIT = False
    
    Dim cmd As Object
    Set cmd = wsh.Exec(exe_path)
    
    '�v���Z�X�������ɒʒm���󂯎��
    Do While cmd.Status = 0
      DoEvents
    Loop
    
    '�v���Z�X�̖߂�l���擾����
    RunProcessWait = cmd.ExitCode
    
    Dim stdout As String
    
    
    Dim str As String
    Dim bytes() As Byte

    str = "��"
    bytes = StrConv(str, vbFromUnicode) '82 A0 (SJIS)
    bytes = StrConv(str, vbUnicode)     '30 42 (Unicode = UTF16/UTF8)
    
                                '46d5494 ��
    str = cmd.stdout.ReadAll    '46d5494 ぁE
    
                                        ' 4  6  d  5  4  9  4
    bytes = StrConv(str, vbFromUnicode) '34 36 64 35 34 39 34 20 E3 81 81 45
    bytes = StrConv(str, vbUnicode)     '34 36 64 35 34 39 34 20 3A 7E FB 30
    
    stdout = ""
    
    'stdout = cmd.stdout.ReadAll
    
    
    'testsub3 cmd
    'test1
    'stdout = TestSub1(cmd)
    'stdout = ReadStdOut(cmd)
    'TestSub2

    Set cmd = Nothing
    Set wsh = Nothing
End Function

Sub testsub5()
    Const git As String = """C:\Program Files\Git\cmd\git.exe"""

    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim ex As Object

    Dim cmd  ' ���s�R�}���h
    Dim aryCmd(2)  ' ���s�R�}���h�z��
    Dim gitCmd  ' git�R�}���h
    Dim aryGitCmd(1)  ' git�R�}���h�z��
    Dim result As String  ' �R�}���h���s����

    '// git�R�}���h��z��Ɋi�[
    aryGitCmd(0) = git
    aryGitCmd(1) = "log --oneline"

    '// git�R�}���h���󔒋�؂�ŘA��
    gitCmd = Join(aryGitCmd, " ")
    MsgBox "gitCmd > " & gitCmd

    '// ���s���鏇�ɃR�}���h��z��Ɋi�[
    aryCmd(0) = "set LANG=ja_JP.UTF-8"
    aryCmd(1) = "C:"
    aryCmd(2) = gitCmd

    '// �R�}���h��A��
    cmd = Join(aryCmd, " & ")
    MsgBox "cmd > " & cmd

    '// �R�}���h���s
    Set ex = wsh.Exec("cmd.exe /C " & cmd)
    
    '// �R�}���h���s��
    If (ex.Status <> 0) Then
        '// �����𔲂���
        MsgBox "�����Ɏ��s���܂���"
        Exit Sub
    End If

    '// �R�}���h���s���͑҂�
    Do While (ex.Status = 0)
        DoEvents
    Loop

    '// �W���o�͂̌��ʂ�\������
    result = ex.stdout.ReadAll
End Sub

Sub testsub4()
    Dim wsh As Object
    Dim cmd As Object
    Set wsh = CreateObject("WScript.Shell")
    Set cmd = wsh.Exec("ipconfig.exe")
    
    Dim strLine As String
    Do Until cmd.stdout.AtEndOfStream      ' �W���o�͂��I������܂Ń��[�v
      strLine = cmd.stdout.ReadLine         ' 1�s�ǂݍ���
    Loop
End Sub

Sub testsub3(ByRef cmd As Object)
    Dim strLine As String
    
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    'Const ENCODE = "UTF-8"
    Const ENCODE = "shift_jis"

    
    With adoSt
        .Charset = ENCODE
        .Type = 2
        .LineSeparator = -1
        .Open
    End With

    Do Until cmd.stdout.AtEndOfStream
        
        strLine = cmd.stdout.ReadLine
        
        With adoSt
            .WriteText strLine, 1
        End With

        Debug.Print strLine
    Loop
    
    With adoSt
        .SaveToFile "C:\_tmp\test.txt", 2
        .Close
    End With
    
    Set adoSt = Nothing
End Sub

Sub TestSub2()
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    ' �����ݒ�
    With adoSt
        .Charset = "UTF-8" ' ���ꂾ������ UTF-8BOM�t�� �ɂȂ�c
        .Type = 2
        .LineSeparator = -1
    End With
    
    ' ��������
    With adoSt
        .Open
        .WriteText "abc", 1
        .WriteText "1234", 1
        .WriteText "����������", 1
        .SaveToFile "C:\_tmp\UTF-8BOM�t��.txt", 2
        .Close
    End With
    
    Set adoSt = Nothing
End Sub

Private Function TestSub1(ByRef cmd As Object) As String
    Const TYPE_TEXT = 2
    Const OPT_WRITE_LINE = 1
    Const OPT_OVER_WRITE = 2
    Const CRLF = -1
    Const CR = 13
    Const LF = 10
    'Const ENCODE = "UTF-8"
    Const ENCODE = "shift_jis"

    'Dim stdout As String
    'stdout = cmd.stdout.ReadAll
    
    Dim i As Long
    Dim j As Long
    Dim strList As String
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    adoSt.Type = TYPE_TEXT
    adoSt.Charset = ENCODE
    adoSt.LineSeparator = CRLF
    adoSt.Open
    adoSt.WriteText cmd.stdout.ReadAll, OPT_WRITE_LINE
    adoSt.SaveToFile "C:\_tmp\test.txt", OPT_OVER_WRITE
    adoSt.Close
    Set adoSt = Nothing
    
    Dim a As String
    a = Common.ReadTextFileBySJIS("C:\_tmp\test.txt")
    
    
    'Dim s As String: s = StrConv(stdout, vbFromUnicode)
    's = StrConv(s, vbUnicode)
  
    'Dim objStream As Object
    'Set objStream = CreateObject("ADODB.Stream")
    'objStream.Charset = "UTF-8"
    'objStream.Charset = "shift_jis"
    'objStream.Type = 2 ' �e�L�X�g���[�h
    'objStream.LineSeparator = -1 'CRLF
    
    'objStream.Open
    'objStream.WriteText cmd.stdout.ReadAll
    
    'objStream.Position = 0
    'stdout = objStream.ReadText(-1)
    
        ' �^�C�v���o�C�i���ɂ��āA�擪��3�o�C�g���X�L�b�v
    'objStream.Position = 0
    'objStream.Type = 1 ' �^�C�v�ύX����ɂ�Position = 0�ł���K�v������
    'objStream.Position = 3
    ' �ꎞ�i�[�p
    'Dim p_byteData() As Byte
    'p_byteData = objStream.Read
    'objStream.Close ' ��U����
    'objStream.Open ' �ēx�J����
    'objStream.Write p_byteData ' �X�g���[���ɏ�������

    ' ---------- �����܂� ��ǉ� ----------
    
    'objStream.SaveToFile "C:\_tmp\UTF-8BOM�Ȃ�.txt", 2
    'objStream.Close
End Function

Function ReadStdOut(cmd As Object) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' �e�L�X�g���[�h
    'stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText cmd.stdout.ReadAll, 1
    stream.Position = 0
    'stream.Open
    'stream.Charset = "UTF-8"
    Dim utf16Log As String
    utf16Log = stream.ReadText(-1)
    utf16Log = Replace(utf16Log, vbLf, vbCrLf)
    stream.Close
    
    
    ReadStdOut = utf16Log
End Function
