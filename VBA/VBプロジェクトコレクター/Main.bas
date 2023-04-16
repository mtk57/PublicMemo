Attribute VB_Name = "Main"
Option Explicit

'--------------------------------------------------------
'他のブックから呼び出す場合はこのメソッドを使うこと
' vbprj_files : IN : VBプロジェクトファイルパスリスト(絶対パス)
' dst_dir_path : IN : コピー先フォルダパス(絶対パス)
' is_debug : IN : デバッグログ出力有無(True=出力する)
' Ret : True/False (True=成功)
'--------------------------------------------------------
Public Function Run( _
    ByRef vbprj_files() As String, _
    ByVal dst_dir_path As String, _
    ByVal is_debug As Boolean _
    ) As Boolean
    
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Process.IS_EXTERNAL = True
    
    Dim msg As String: msg = "正常に終了しました"
    Dim ret As Boolean: ret = True
    
    If is_debug = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VBPrjCollector.log"
    End If

    '開始
    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"
    
    CreateParamForExternal vbprj_files, dst_dir_path
    Process.Run

    Common.WriteLog "★End"
    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"
    ret = False

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    Run = ret
End Function

'--------------------------------------------------------
'--------------------------------------------------------
Public Sub Run_Click()

On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Process.IS_EXTERNAL = False

    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VBPrjCollector.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("main").Activate
    Process.Run

    Common.WriteLog "★End"
    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "O8"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(CLM).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

Private Sub CreateParamForExternal( _
    ByRef vbprj_files() As String, _
    ByVal dst_dir_path As String _
    )
    Common.WriteLog "CreateParamForExternal S"
    
    Dim main_param As MainParam
    Dim sub_param As SubParam
    Set main_param = New MainParam
    Set sub_param = New SubParam
    
    main_param.InitForExternal dst_dir_path
    sub_param.InitForExternal vbprj_files
    
    Set Process.main_param = main_param
    Set Process.sub_param = sub_param
    
    Common.WriteLog "CreateParamForExternal E"
End Sub
