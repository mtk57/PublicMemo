Attribute VB_Name = "Common"
Option Explicit

'ログファイル番号
Private logfile_num As Integer
Private is_log_opened As Boolean

'-------------------------------------------------------------
'ログファイルをオープンする
' logfile_path : IN : ログファイルパス(絶対パス)
'-------------------------------------------------------------
Public Sub OpenLog(ByVal logfile_path As String)
    If is_log_opened = True Then
        'すでにオープンしているので無視
        Exit Sub
    End If
    logfile_num = FreeFile()
    Open logfile_path For Append As logfile_num
    is_log_opened = True
End Sub

'-------------------------------------------------------------
'ログファイルに書き込む
' contents : IN : 書き込む内容
'-------------------------------------------------------------
Public Sub WriteLog(ByVal contents As String)
    If is_log_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
End Sub

'-------------------------------------------------------------
'ログファイルをクローズする
'-------------------------------------------------------------
Public Sub CloseLog()
    If is_log_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Close logfile_num
    logfile_num = -1
    is_log_opened = False
End Sub

'-------------------------------------------------------------
'ログファイルに書き込む
' contents : IN : 書き込む内容
'-------------------------------------------------------------
Public Sub WriteLogSimple(ByVal contents As String)
    Dim file_num As Integer
    file_num = FreeFile()
    Open "Logger.log" For Append As file_num
    Print #file_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
    Close file_num
    file_num = -1
End Sub

