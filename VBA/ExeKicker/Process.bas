Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String

'パラメータ
Private main_param As MainParam
Private sub_params() As SubParam

'グローバル
Private current_wk_src_dir_path As String
Private current_wk_dst_dir_path As String
Private before_wk_dst_dir_path As String

'メイン処理
Public Sub Run()
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    'パラメータのチェックと収集を行う
    If CheckAndCollectParam() = False Then
        Exit Sub
    End If
    
    'Sub Paramを順に実行していく
    If ExecSubParam() = False Then
        Exit Sub
    End If
    
    MsgBox "終わりました"
End Sub

'作業用フォルダ削除
Public Sub DelWkDir()
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    'パラメータのチェックと収集を行う
    If CheckAndCollectParam() = False Then
        Exit Sub
    End If
    
    DeleteWorkFolder True
    
    MsgBox "終わりました"
End Sub

'パラメータのチェックと収集を行う
Private Function CheckAndCollectParam() As Boolean
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    err_msg = main_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        CheckAndCollectParam = False
        Exit Function
    End If

    
    'Sub Params
    Const START_ROW = 21
    Const SUB_ROWS = 5
    Dim row As Integer: row = START_ROW
    Dim cnt As Integer: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        err_msg = sub_param.Init(row)
        If err_msg <> "" Then
            MsgBox err_msg
            CheckAndCollectParam = False
            Exit Function
        End If
        
        If sub_param.GetEnable() = "STOPPER" Then
            Exit Do
        ElseIf sub_param.GetEnable() = "DISABLE" Then
            GoTo CONTINUE
        End If
        
        ReDim Preserve sub_params(cnt)
        Set sub_params(cnt) = sub_param
        cnt = cnt + 1
        
CONTINUE:
        row = row + SUB_ROWS + 1
    Loop

    CheckAndCollectParam = True
End Function

'Sub Paramを順に実行していく
Private Function ExecSubParam() As Boolean
    If UBound(sub_params) < 0 Then
        MsgBox "有効なSub paramがありません"
        ExecSubParam = True
        Exit Function
    End If

    Dim i, j As Integer
    Dim exe_params() As String
    Dim is_match As Boolean
    Dim is_exit_for As Boolean
    
    '作業用フォルダを作成する
    CreateWorkFolder
    
    '除外リストファイルを作成する
    CreateIgnoreListFile
    
    For i = LBound(sub_params) To UBound(sub_params)
        is_exit_for = False
        
        Dim sub_param As SubParam
        Set sub_param = sub_params(i)
        
        'exeiniを更新する
        UpdateExeIniContents sub_param
        
        For j = 0 To main_param.GetMaxExecCount() - 1
        
            '連番の作業用サブフォルダを作成する
            CreateSeqWorkFolder i, j
            
            '作業用サブフォルダのsrc→dstにコピーする
            CopySrcToDstWorkFolder i, j
            
            'exeに渡すパラメータリストを作成する
            exe_params = CreateExeParamList(sub_param)
            
            'exeを実行する
            RunExe exe_params
            
            'TODO:除外ファイルフォーマット不明のままの場合、ここでsrc→dstに除外ファイルをコピーする
            
            '差分があるかチェックする
            is_match = IsMatch()
            If is_match = True Then
                '全て一致
                is_exit_for = True
            Else
                '1つ以上の不一致あり
                If sub_param.IsExecNotDiff() = False Then
                    is_exit_for = True
                End If
            End If
            
            before_wk_dst_dir_path = current_wk_dst_dir_path
            
            '作業用サブフォルダを入れ替える
            SwapWorkSubFolder
            
            If is_exit_for = True Then
                Exit For
            End If

        Next j
    
    Next i
    
    'dstにコピーする
    If main_param.IsStepWorkDir() = False Then
        current_wk_dst_dir_path = before_wk_dst_dir_path
    End If
    
    Common.CopyFolder current_wk_dst_dir_path, main_param.GetDestDirPath
    
    '作業用フォルダを削除する
    DeleteWorkFolder main_param.IsDeleteWorkDir()

    ExecSubParam = True
End Function

'作業用フォルダを作成する
Private Sub CreateWorkFolder()
    Dim path As String: path = main_param.GetToolWorkDirPath()

    If path = "" Then
        '未指定の場合はC:\tmpとする
        path = "C:\tmp"
        main_param.SetToolWorkDirPath (path)
    End If

    Common.CreateFolder (path)
    
    If main_param.IsStepWorkDir() = False Then
        '途中経過残さない場合、固定サブフォルダを作成
        current_wk_src_dir_path = path & SEP & "FIX" & "_0"
        current_wk_dst_dir_path = path & SEP & "FIX" & "_1"
        Common.CreateFolder (current_wk_src_dir_path)
        Common.CreateFolder (current_wk_dst_dir_path)
    End If
End Sub

'除外リストファイルを作成する
Private Sub CreateIgnoreListFile()
    If UBound(main_param.GetIgnoreFiles()) < 0 Then
        '除外ファイルなし
        main_param.SetIgnoreFilePath ("")
        Exit Sub
    End If
    
    '除外リストファイルパス
    Const IGNORE_FILE_NAME = "TODO.ini"
    Dim path As String: path = main_param.GetToolWorkDirPath() & SEP & IGNORE_FILE_NAME
    
    main_param.SetIgnoreFilePath (path)
    
    '除外リストファイルを作成
    Dim filelist() As String: filelist = main_param.GetIgnoreFiles()
    Dim i As Integer
    For i = LBound(filelist) To UBound(filelist)
        Dim num As Integer: num = i + 1
        Dim ret As Integer: ret = Common.WritePrivateProfileString("TODO", "Name" & num, filelist(i), path)
        If ret = 0 Then
            Err.Raise 53, , "除外リストファイルの更新に失敗しました"
        End If
    Next i
End Sub

'exeiniを更新
Private Sub UpdateExeIniContents(ByRef sub_param As SubParam)
    Dim ret As Long
    Dim path As String: path = main_param.GetExeIniFilePath()
    
    Dim addin_path As String: addin_path = main_param.GetAddinFilePath()
    ret = Common.WritePrivateProfileString("Extent", "Name1", addin_path, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(0)"
    End If

    Dim count As String
    If sub_param.IsEnableAddin = True Then
        count = "1"
    Else
        count = "0"
    End If
    ret = Common.WritePrivateProfileString("Extent", "Count", count, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(1)"
    End If
    
    Dim skip As String
    If sub_param.IsSkipComment = True Then
        skip = "1"
    Else
        skip = "0"
    End If
    ret = Common.WritePrivateProfileString("Comment", "SkipVb", skip, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(2)"
    End If
End Sub

'連番の作業用サブフォルダを作成する
Private Sub CreateSeqWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    If main_param.IsStepWorkDir() = False Then
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    current_wk_src_dir_path = path & SEP & num1 & num2 & "_0"
    current_wk_dst_dir_path = path & SEP & num1 & num2 & "_1"
    Common.CreateFolder (current_wk_src_dir_path)
    Common.CreateFolder (current_wk_dst_dir_path)
End Sub

'作業用サブフォルダのsrc→dstにコピーする
Private Sub CopySrcToDstWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer)
    Dim src_path As String
    Dim dst_path_0 As String
    Dim dst_path_1 As String
    
    If main_param.IsStepWorkDir() = True Then
        '途中経過を残す
        
        If num1 = 0 And num2 = 0 Then
            '最初だけは本当のsrcからコピーする
            src_path = main_param.GetSrcDirPath()
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
        Else
            src_path = before_wk_dst_dir_path
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
        End If
        
        Common.CopyFolder src_path, dst_path_0
        Common.CopyFolder src_path, dst_path_1
    
    Else
        '途中経過を残さない
        
        If num1 = 0 And num2 = 0 Then
            '最初だけは本当のsrcからコピーする
            src_path = main_param.GetSrcDirPath()
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
            
            Common.CopyFolder src_path, dst_path_0
            Common.CopyFolder src_path, dst_path_1
        Else
            Common.CopyFolder current_wk_src_dir_path, current_wk_dst_dir_path
        End If
        
    End If
        
End Sub

'exeに渡すパラメータリスト作成する
Private Function CreateExeParamList(ByRef sub_param As SubParam) As String()
    Dim i As Integer
    Dim param_list() As String
    
    Dim src_path_list() As String
    Dim dst_path_list() As String
    
    If main_param.IsContainSubDir() = False Then
        ReDim src_path_list(0)
        ReDim dst_path_list(0)
        src_path_list(0) = current_wk_src_dir_path
        dst_path_list(0) = current_wk_dst_dir_path
    Else
        src_path_list = Common.GetFolderPathList(current_wk_src_dir_path)
        dst_path_list = Common.GetFolderPathList(current_wk_dst_dir_path)
        
        Common.AppendArray src_path_list, current_wk_src_dir_path
        Common.AppendArray dst_path_list, current_wk_dst_dir_path
    End If
    
    For i = LBound(src_path_list) To UBound(src_path_list)
        ReDim Preserve param_list(i)
        '"srcdirpath" "dstdirpath" "*.vb" "inipath" "ignorefilelistpath" ""
        param_list(i) = _
            Chr(34) & src_path_list(i) & Chr(34) & " " & _
            Chr(34) & dst_path_list(i) & Chr(34) & " " & _
            Chr(34) & main_param.GetInExtension() & Chr(34) & " " & _
            Chr(34) & sub_param.GetIniFilePath() & Chr(34) & " " & _
            Chr(34) & main_param.GetIgnoreFilePath() & Chr(34) & " " & _
            Chr(34) & Chr(34)
    Next i
    
    CreateExeParamList = param_list
End Function

'exeを実行する
Private Sub RunExe(ByRef param_list() As String)
    Dim i As Integer
    Dim ret As Long
    Dim exe_param As String
    
    For i = LBound(param_list) To UBound(param_list)
        
        exe_param = _
            Chr(34) & main_param.GetExeFilePath() & Chr(34) & " " & _
            param_list(i)
        
        ret = Common.RunProcessWait(exe_param)
        
        If ret <> 0 Then
            Err.Raise 53, , "Exeの実行に失敗しました(ret=" & ret & ")"
        End If
    
    Next i

End Sub

'差分があるかチェックする
Private Function IsMatch() As Boolean
    Dim i As Integer
    Dim is_match As Boolean: is_match = True

    'ファイルリストを作成
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(current_wk_src_dir_path, main_param.GetInExtension())
    Dim dst_file_list() As String: dst_file_list = Common.CreateFileList(current_wk_dst_dir_path, main_param.GetInExtension())

    'ファイルを比較
    For i = LBound(src_file_list) To UBound(src_file_list)
        is_match = Common.IsMatchTextFiles(src_file_list(i), dst_file_list(i))
        If is_match = False Then
            '1つでも差異があれば終了
            IsMatch = is_match
            Exit Function
        End If
    Next i
    
    IsMatch = is_match
End Function

'作業用サブフォルダを入れ替える
Private Sub SwapWorkSubFolder()
    If main_param.IsStepWorkDir() = True Then
        Exit Sub
    End If
    
    Dim tmp As String: tmp = current_wk_src_dir_path
    current_wk_src_dir_path = current_wk_dst_dir_path
    current_wk_dst_dir_path = tmp
End Sub

'作業用フォルダを削除する
Private Sub DeleteWorkFolder(ByVal is_del_wk_dir As Boolean)
    If is_del_wk_dir = False Then
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    
    If Common.IsExistsFolder(path) = False Then
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("作業用フォルダを削除しますか?") = False Then
        Exit Sub
    End If
    
    Common.DeleteFolder path
End Sub
