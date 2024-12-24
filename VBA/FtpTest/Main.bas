Attribute VB_Name = "Main"
#If VBA7 Then ' VBA7以降 (Office 2010以降) の場合

  ' 64bit互換のWinExec宣言
  Declare PtrSafe Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As LongPtr) As LongPtr

#Else ' VBA6以前 (Office 2007以前) の場合 (互換性のために残す)

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

  ' FTPサーバー情報
  Const FTP_SERVER As String = "localhost"
  Const FTP_USER As String = "ftpuser"
  Const FTP_PASSWORD As String = "ftpuser"
  Const REMOTE_FILE As String = "/test.txt"   ' ダウンロードするリモートファイルパス
  Const LOCAL_FILE As String = "C:\_git\PublicMemo\VBA\FtpTest\local_test.txt"

  Dim scriptPath As String
  Dim scriptContent As String
  Dim cmd As String

  Dim ftpPath As String
  Dim result As LongPtr ' LongPtr型に変更

  ' ftp.exeのフルパスを取得
  ftpPath = Environ("SystemRoot") & "\System32\ftp.exe"

  ' スクリプトファイルパスを作成 (一時ファイルとして作成)
  scriptPath = Environ("TEMP") & "\ftp_script.txt"

  ' スクリプト内容を作成
  scriptContent = "open " & FTP_SERVER & vbCrLf & _
                  FTP_USER & vbCrLf & _
                  FTP_PASSWORD & vbCrLf & _
                  "ascii" & vbCrLf & _
                  "get " & REMOTE_FILE & " " & LOCAL_FILE & vbCrLf & _
                  "bye"

  ' スクリプトファイルを保存
  Open scriptPath For Output As #1
    Print #1, scriptContent
  Close #1

  ' ftp.exeを実行 (パスにスペースが含まれる場合を考慮して""で囲む)
  cmd = """" & ftpPath & """ -s:" & scriptPath

  ' カレントディレクトリを一時フォルダに設定 (重要)
  ChDir Environ("TEMP")

  ' WinExecで実行
  result = WinExec(cmd, 0) ' 0 は vbHide と同じ

  Select Case result
    Case 0: MsgBox "FTPコマンドの実行に失敗しました。ftp.exeが見つからない可能性があります。", vbCritical
    Case 1 To 31: MsgBox "FTPコマンド実行中にエラーが発生しました。エラーコード：" & result, vbCritical
    Case Else: MsgBox "FTP処理が完了しました。", vbInformation
  End Select

  ' (オプション) スクリプトファイルを削除 (必要に応じてコメントアウト)
  Kill scriptPath

End Sub



