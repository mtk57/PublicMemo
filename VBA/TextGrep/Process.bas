Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Public main_param As MainParam
Private sub_params() As SubParam

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Erase sub_params

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    'grep.exeに渡すファイルを作成する
    main_param.CreateInputFile sub_params
    
    'grep.exeを実行する
    RunExe
    
    '結果ファイルを読み込みシートに出力する
    OutputSheet

    Common.WriteLog "Run E"
End Sub

'パラメータのチェックと収集を行う
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    Set main_param = New MainParam
    main_param.Init
    main_param.Validate
    
    Dim i As Long: i = 0
    Dim row As Long: row = 18
    Dim cnt As Long: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        sub_param.Init row + i
        
        If sub_param.GetKeyword = "" Then
            Exit Do
        End If
        
        ReDim Preserve sub_params(cnt)
        Set sub_params(cnt) = sub_param
        cnt = cnt + 1
        
        i = i + 1
    Loop

    Common.WriteLog "CheckAndCollectParam E"
End Sub

Private Sub RunExe()
    Common.WriteLog "RunExe S"

    Dim ret As Long
    Dim exe_param As String
    
    exe_param = _
        DQ & _
        main_param.GetGrepExe() & _
        DQ & _
        " " & _
        DQ & _
        main_param.GetInputFilePath() & _
        DQ
    
    Common.WriteLog exe_param
    
    ret = Common.RunProcessWait(exe_param)
        
    If ret <> 0 Then
        Err.Raise 53, , "Exeの実行に失敗しました(exe ret=" & ret & ")"
    End If
    
    Common.WriteLog "RunExe E"
End Sub

Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
    
    'Common.AddSheet ActiveWorkbook, main_param.GetOutputSheetName()
    
    Workbooks.OpenText FILENAME:=main_param.GetOutputFilePath(), _
        DataType:=xlDelimited, _
        Tab:=True

    Dim ws As Worksheet
    Set ws = Workbooks.Item(Workbooks.count).Sheets(1)                  ' 開いた CSV のシート
    ws.name = main_param.GetOutputSheetName()
    Call ws.Move(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)) ' VBA を実行しているワークブックの最後に移動
    
    Common.WriteLog "OutputSheet E"
End Sub

