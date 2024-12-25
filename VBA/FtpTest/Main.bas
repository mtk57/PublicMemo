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

'PowerShell & 環境変数版
Private Sub ExecuteFTPScript_2()

    ' 一時的な環境変数名
    Const TEMP_USER_VAR As String = "TEMP_FTP_USER"
    Const TEMP_PASSWORD_VAR As String = "TEMP_FTP_PASSWORD"
    Const TEMP_SERVER_VAR As String = "TEMP_FTP_SERVER"
    Const TEMP_REMOTE_FILE_VAR As String = "TEMP_FTP_REMOTE_FILE"
    Const TEMP_LOCAL_FILE_VAR As String = "TEMP_FTP_LOCAL_FILE"

    ' FTP情報
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
    
    ' 環境変数を設定
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

    ' PowerShellスクリプトの内容をVBAで生成
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
        "  Write-Host ""FTPファイルのダウンロードが完了しました。""" & vbCrLf & _
        "} catch {" & vbCrLf & _
        "  Write-Error ""FTPファイルのダウンロードに失敗しました: $($_.Exception.Message)""" & vbCrLf & _
        "  exit 1" & vbCrLf & _
        "}"


    ' PowerShellスクリプトを一時ファイルに保存
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\ftp_script.ps1"
       
    Open scriptPath For Output As #1
      Print #1, scriptContent
    Close #1
    
    Dim cmd As String
    cmd = "powershell.exe -ExecutionPolicy Bypass -File """ & scriptPath & """"

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    result = objShell.Run(cmd, 0, True) ' 0: 非表示, 1: 通常表示, True: 完了まで待機

    If result = 0 Then
        MsgBox "PowerShellスクリプトが正常に完了しました。", vbInformation
    Else
        MsgBox "PowerShellスクリプトの実行に失敗しました。エラーコード: " & result, vbCritical
    End If

    Set objShell = Nothing
    
    ' 環境変数を削除
    SetEnvironmentVariable TEMP_USER_VAR, vbNullString
    SetEnvironmentVariable TEMP_PASSWORD_VAR, vbNullString
    SetEnvironmentVariable TEMP_SERVER_VAR, vbNullString
    SetEnvironmentVariable TEMP_REMOTE_FILE_VAR, vbNullString
    SetEnvironmentVariable TEMP_LOCAL_FILE_VAR, vbNullString
    
    ' スクリプトファイルを削除
    Kill scriptPath

End Sub



