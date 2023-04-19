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
        Err.Raise 53, , "Exeの実行に失敗しました(ret=" & ret & ")"
    End If

    Common.WriteLog "RunExe E"
End Sub

'-------------------------------------------------------------
'外部アプリケーションを実行し、終了するまで待機する
' exe_path : IN : 外部アプリケーション(exe)の絶対パス
'                 exeに渡すパラメータがある場合も一緒に書くこと
' Ret : プロセスの戻り値
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
    
    'プロセス完了時に通知を受け取る
    Do While cmd.Status = 0
      DoEvents
    Loop
    
    'プロセスの戻り値を取得する
    RunProcessWait = cmd.ExitCode
    
    Dim stdout As String
    
    
    Dim str As String
    Dim bytes() As Byte

    str = "あ"
    bytes = StrConv(str, vbFromUnicode) '82 A0 (SJIS)
    bytes = StrConv(str, vbUnicode)     '30 42 (Unicode = UTF16/UTF8)
    
                                '46d5494 あ
    str = cmd.stdout.ReadAll    '46d5494 縺・
    
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

    Dim cmd  ' 実行コマンド
    Dim aryCmd(2)  ' 実行コマンド配列
    Dim gitCmd  ' gitコマンド
    Dim aryGitCmd(1)  ' gitコマンド配列
    Dim result As String  ' コマンド実行結果

    '// gitコマンドを配列に格納
    aryGitCmd(0) = git
    aryGitCmd(1) = "log --oneline"

    '// gitコマンドを空白区切りで連結
    gitCmd = Join(aryGitCmd, " ")
    MsgBox "gitCmd > " & gitCmd

    '// 実行する順にコマンドを配列に格納
    aryCmd(0) = "set LANG=ja_JP.UTF-8"
    aryCmd(1) = "C:"
    aryCmd(2) = gitCmd

    '// コマンドを連結
    cmd = Join(aryCmd, " & ")
    MsgBox "cmd > " & cmd

    '// コマンド実行
    Set ex = wsh.Exec("cmd.exe /C " & cmd)
    
    '// コマンド失敗時
    If (ex.Status <> 0) Then
        '// 処理を抜ける
        MsgBox "処理に失敗しました"
        Exit Sub
    End If

    '// コマンド実行中は待ち
    Do While (ex.Status = 0)
        DoEvents
    Loop

    '// 標準出力の結果を表示する
    result = ex.stdout.ReadAll
End Sub

Sub testsub4()
    Dim wsh As Object
    Dim cmd As Object
    Set wsh = CreateObject("WScript.Shell")
    Set cmd = wsh.Exec("ipconfig.exe")
    
    Dim strLine As String
    Do Until cmd.stdout.AtEndOfStream      ' 標準出力が終了するまでループ
      strLine = cmd.stdout.ReadLine         ' 1行読み込み
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
    
    ' 初期設定
    With adoSt
        .Charset = "UTF-8" ' これだけだと UTF-8BOM付き になる…
        .Type = 2
        .LineSeparator = -1
    End With
    
    ' 書き込み
    With adoSt
        .Open
        .WriteText "abc", 1
        .WriteText "1234", 1
        .WriteText "あいうえお", 1
        .SaveToFile "C:\_tmp\UTF-8BOM付き.txt", 2
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
    'objStream.Type = 2 ' テキストモード
    'objStream.LineSeparator = -1 'CRLF
    
    'objStream.Open
    'objStream.WriteText cmd.stdout.ReadAll
    
    'objStream.Position = 0
    'stdout = objStream.ReadText(-1)
    
        ' タイプをバイナリにして、先頭の3バイトをスキップ
    'objStream.Position = 0
    'objStream.Type = 1 ' タイプ変更するにはPosition = 0である必要がある
    'objStream.Position = 3
    ' 一時格納用
    'Dim p_byteData() As Byte
    'p_byteData = objStream.Read
    'objStream.Close ' 一旦閉じて
    'objStream.Open ' 再度開いて
    'objStream.Write p_byteData ' ストリームに書き込む

    ' ---------- ここまで を追加 ----------
    
    'objStream.SaveToFile "C:\_tmp\UTF-8BOMなし.txt", 2
    'objStream.Close
End Function

Function ReadStdOut(cmd As Object) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' テキストモード
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
