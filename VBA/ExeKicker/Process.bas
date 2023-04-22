Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Private main_param As MainParam
Private sub_params() As SubParam

'グローバル
Private current_wk_src_dir_path As String
Private current_wk_dst_dir_path As String
Private before_wk_dst_dir_path As String

'メイン処理
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    'Sub Paramを順に実行していく
    ExecSubParam
    
    Common.WriteLog "Run E"
End Sub

'作業用フォルダ削除
Public Sub DelWkDir()
    Common.WriteLog "DelWkDir S"
    
    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    DeleteWorkFolder True
    
    Common.WriteLog "DelWkDir E"
End Sub

'パラメータのチェックと収集を行う
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    main_param.Init
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()

    'Sub Params
    Const START_ROW = 22
    Const SUB_ROWS = 5
    Dim row As Integer: row = START_ROW
    Dim cnt As Integer: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        Common.WriteLog "row=" & row
        sub_param.Init row
        sub_param.Validate

        Common.WriteLog sub_param.GetAllValue()
        
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

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'Sub Paramを順に実行していく
Private Sub ExecSubParam()
    Common.WriteLog "ExecSubParam S"
    
    Dim errmsg As String
    
    If UBound(sub_params) < 0 Then
        Err.Raise 53, , "有効なSub paramがありません"
    End If

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim exe_params() As String
    Dim is_match As Boolean
    Dim is_exit_for As Boolean
    
    '対象拡張子のファイルが存在するか確認する
    Dim ext As String: ext = Replace(main_param.GetInExtension(), "*", "")
    If Common.IsExistsExtensionFile(main_param.GetSrcDirPath(), ext) = False Then
        Err.Raise 53, , "処理対象の拡張子のファイルが存在しません (" & ext & ")"
    End If
    
    '作業用フォルダを作成する
    CreateWorkFolder
    
    '除外リストファイルを作成する
    CreateIgnoreListFile

    
    For i = LBound(sub_params) To UBound(sub_params)
    
        Dim sub_param As SubParam
        Set sub_param = sub_params(i)
        
        'SJIS, UTF8ごとに実行する(0=SJIS, 1=UTF8)
        For h = 0 To 1
        
            Dim unmatch_cnt As Integer: unmatch_cnt = 0
        
            'exeiniを更新する
            UpdateExeIniContents sub_param, h
            
            For j = 0 To main_param.GetMaxExecCount() - 1
            
                Common.WriteLog "[■Main Loop] i=" & i & ", h=" & h & ", j=" & j
            
                '連番の作業用サブフォルダを作成する
                CreateSeqWorkFolder i, h, j
                
                '作業用サブフォルダのsrc→dstにコピーする
                CopySrcToDstWorkFolder i, j, h
                
                'exeに渡すパラメータリストを作成する
                exe_params = CreateExeParamList(sub_param)
                
                'exeを実行する
                RunExe exe_params
                
                '差分があるかチェックする
                is_exit_for = False
                is_match = IsMatch()
                If is_match = True Then
                    '全て一致
                    is_exit_for = True
                Else
                    '1つ以上の不一致あり
                    If sub_param.IsExecNotDiff() = False Then
                        is_exit_for = True
                    Else
                        unmatch_cnt = unmatch_cnt + 1
                    End If
                End If
                
                before_wk_dst_dir_path = current_wk_dst_dir_path
                
                '作業用サブフォルダを入れ替える
                SwapWorkSubFolder
                
                If is_exit_for = True Then
                    Exit For
                End If
    
            Next j
            
            If unmatch_cnt >= main_param.GetMaxExecCount() And IsTestMode() = False Then
                Err.Raise 53, , "不一致が最大値以上になったので中止します(" & unmatch_cnt & ")"
            End If
        
        Next h
    
    Next i
    

    If main_param.IsStepWorkDir() = False Then
        current_wk_dst_dir_path = before_wk_dst_dir_path
    End If
    
    '最後に作業フォルダのリネームされた拡張子を元に戻す
    RenameAllFileExtension current_wk_dst_dir_path, -1
    
    'dstにコピーする
    Common.DeleteFolder (main_param.GetDestDirPath())
    Common.CopyFolder current_wk_dst_dir_path, main_param.GetDestDirPath()
    
    '作業用フォルダを削除する
    DeleteWorkFolder main_param.IsDeleteWorkDir()

    Common.WriteLog "ExecSubParam E"
End Sub

'作業用フォルダを作成する
Private Sub CreateWorkFolder()
    Common.WriteLog "CreateWorkFolder S"

    Dim path As String: path = main_param.GetToolWorkDirPath()

    If path = "" Then
        '未指定の場合はC:\tmpとする
        path = "C:\tmp"
        main_param.SetToolWorkDirPath (path)
    End If
    
    Common.DeleteFolder (path)

    Common.CreateFolder (path)
    
    If main_param.IsStepWorkDir() = False Then
        '途中経過残さない場合、固定サブフォルダを作成
        current_wk_src_dir_path = path & SEP & "FIX" & "_0"
        current_wk_dst_dir_path = path & SEP & "FIX" & "_1"
        Common.CreateFolder (current_wk_src_dir_path)
        Common.CreateFolder (current_wk_dst_dir_path)
    End If
    
    Common.WriteLog "CreateWorkFolder E"
End Sub

'除外リストファイルを作成する
Private Sub CreateIgnoreListFile()
    Common.WriteLog "CreateIgnoreListFile S"

    If UBound(main_param.GetIgnoreFiles()) < 0 Then
        '除外ファイルなし
        main_param.SetIgnoreFilePath ("")
        
        Common.WriteLog "CreateIgnoreListFile E1"
        Exit Sub
    End If
    
    '除外リストファイルパス
    Const IGNORE_FILE_NAME = "ExclusionList.ini"
    Dim path As String: path = main_param.GetToolWorkDirPath() & SEP & IGNORE_FILE_NAME
    
    main_param.SetIgnoreFilePath (path)
    
    '除外リストファイルを作成
    Dim filelist() As String: filelist = main_param.GetIgnoreFiles()
    Dim i As Integer
    For i = LBound(filelist) To UBound(filelist)
        Dim num As Integer: num = i + 1
        Dim ret As Integer: ret = Common.WritePrivateProfileString("SkipFile", "File" & num, filelist(i), path)
        If ret = 0 Then
            Err.Raise 53, , "除外リストファイルの更新に失敗しました"
        End If
    Next i
    
    Common.WriteLog "CreateIgnoreListFile E"
End Sub

'全ファイルを変換されないように拡張子をリネーム
Private Sub RenameAllFileExtension(ByVal path As String, ByVal encode_type As Integer)
    Common.WriteLog "RenameAllFileExtension S"
    Common.WriteLog "path=(" & path & "), encode_type=(" & encode_type & ")"
    
    Dim all_file_list() As String
    Dim i As Long
    Dim renamed_path As String
    
    Const EXT_UTF8 = ".utf8"
    Const EXT_SJIS = ".sjis"
    
    'ファイルリストを作成
    all_file_list = Common.CreateFileList(path, main_param.GetInExtension(), True)
    
    If encode_type = 0 Then
        '全てのファイルを読み込み、SJIS以外であれば拡張子をリネームする
        For i = 0 To UBound(all_file_list)
            If Common.IsSJIS(all_file_list(i)) = False Then
                Common.WriteLog "Before Rename(UTF8)=" & all_file_list(i)
            
                renamed_path = Common.ChangeFileExt(all_file_list(i), EXT_UTF8)
                
                Common.WriteLog "After Rename(UTF8)=" & renamed_path
            End If
        Next i
    ElseIf encode_type = 1 Then
        '全てのファイルを読み込み、SJISであれば拡張子をリネームする
        For i = 0 To UBound(all_file_list)
            If all_file_list(i) = "" Then
                Exit For
            End If
            
            If Common.IsSJIS(all_file_list(i)) = True Then
                Common.WriteLog "Before Rename(SJIS)=" & all_file_list(i)
            
                renamed_path = Common.ChangeFileExt(all_file_list(i), EXT_SJIS)
                
                Common.WriteLog "After Rename(SJIS)=" & renamed_path
            End If
        Next i
        
        'リネーム済の拡張子は元に戻す(UTF8)
        Dim utf8_file_list() As String
        utf8_file_list = Common.CreateFileList(path, "*" & EXT_UTF8, True)
        
        For i = 0 To UBound(utf8_file_list)
            If utf8_file_list(i) = "" Then
                Exit For
            End If
            
            Common.WriteLog "Before Rename(UTF8)=" & utf8_file_list(i)
        
            renamed_path = Common.ChangeFileExt(utf8_file_list(i), Replace(main_param.GetInExtension(), "*", ""))
            
            Common.WriteLog "After Rename(UTF8)=" & renamed_path
        Next i
    Else
        '最後にリネーム済の拡張子は元に戻す(SJIS)
        Dim sjis_file_list() As String
        sjis_file_list = Common.CreateFileList(path, "*" & EXT_SJIS, True)
        
        For i = 0 To UBound(sjis_file_list)
            If sjis_file_list(i) = "" Then
                Exit For
            End If
            
            Common.WriteLog "Before Rename(SJIS)=" & sjis_file_list(i)
        
            renamed_path = Common.ChangeFileExt(sjis_file_list(i), Replace(main_param.GetInExtension(), "*", ""))
            
            Common.WriteLog "After Rename(SJIS)=" & renamed_path
        Next i
    End If
    
    Common.WriteLog "RenameAllFileExtension E"
End Sub

'exeiniを更新
Private Sub UpdateExeIniContents(ByRef sub_param As SubParam, ByVal encode_type As Integer)
    Common.WriteLog "UpdateExeIniContents S"
    Common.WriteLog "encode_type=(" & encode_type & ")"

    Dim ret As Long
    Dim path As String: path = main_param.GetExeIniFilePath()
    
    Dim addin_path As String: addin_path = main_param.GetAddinFilePath()
    ret = Common.WritePrivateProfileString("Extent", "Name1", addin_path, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(0)"
    End If

    Dim count As String: count = "0"
    If sub_param.IsEnableAddin = True Then
        count = "1"
    End If
    ret = Common.WritePrivateProfileString("Extent", "Count", count, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(1)"
    End If
    
    Dim skip As String: skip = "0"
    If sub_param.IsSkipComment = True Then
        skip = "1"
    End If
    ret = Common.WritePrivateProfileString("Comment", "SkipVb", skip, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(2)"
    End If
    
    Dim encode As String: encode = "0"
    If encode_type = 1 Then
        encode = "1"
    End If
    ret = Common.WritePrivateProfileString("Convart", "baseChar", encode, path)
    If ret = 0 Then
        Err.Raise 53, , "Iniファイルの更新に失敗しました(3)"
    End If
    
    Common.WriteLog "UpdateExeIniContents E"
End Sub

'連番の作業用サブフォルダを作成する
Private Sub CreateSeqWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer, ByVal num3 As Integer)
    Common.WriteLog "CreateSeqWorkFolder S"

    If main_param.IsStepWorkDir() = False Then
        Common.WriteLog "CreateSeqWorkFolder E1"
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    current_wk_src_dir_path = path & SEP & num1 & num2 & num3 & "_0"
    current_wk_dst_dir_path = path & SEP & num1 & num2 & num3 & "_1"
    Common.CreateFolder (current_wk_src_dir_path)
    Common.CreateFolder (current_wk_dst_dir_path)
    
    Common.WriteLog "CreateSeqWorkFolder E"
End Sub

'作業用サブフォルダのsrc→dstにコピーする
Private Sub CopySrcToDstWorkFolder(ByVal num1 As Integer, ByVal num2 As Integer, ByVal encode_type As Integer)
    Common.WriteLog "CopySrcToDstWorkFolder S"

    Dim src_path As String
    Dim dst_path_0 As String
    Dim dst_path_1 As String
    
    If main_param.IsStepWorkDir() = True Then
        '途中経過を残す
        
        If num1 = 0 And num2 = 0 Then
            If encode_type = 0 Then
                '最初だけは本当のsrcからコピーする
                src_path = main_param.GetSrcDirPath()
                dst_path_0 = current_wk_src_dir_path
                dst_path_1 = current_wk_dst_dir_path
            
                Common.CopyFolder src_path, dst_path_0
            
                '全ファイルを変換されないように拡張子をリネーム
                RenameAllFileExtension dst_path_0, encode_type
            
                Common.CopyFolder dst_path_0, dst_path_1
            Else
                src_path = before_wk_dst_dir_path
                dst_path_0 = current_wk_src_dir_path
                dst_path_1 = current_wk_dst_dir_path
                
                Common.CopyFolder src_path, dst_path_0
                
                '全ファイルを変換されないように拡張子をリネーム
                RenameAllFileExtension dst_path_0, encode_type
                
                Common.CopyFolder dst_path_0, dst_path_1
            End If

        Else
            src_path = before_wk_dst_dir_path
            dst_path_0 = current_wk_src_dir_path
            dst_path_1 = current_wk_dst_dir_path
            
            Common.CopyFolder src_path, dst_path_0
            Common.CopyFolder src_path, dst_path_1
        End If
    
    Else
        '途中経過を残さない
        
        If num1 = 0 And num2 = 0 Then
            If encode_type = 0 Then
                '最初だけは本当のsrcからコピーする
                src_path = main_param.GetSrcDirPath()
                dst_path_0 = current_wk_src_dir_path
                dst_path_1 = current_wk_dst_dir_path
                
                Common.CopyFolder src_path, dst_path_0
                
                '全ファイルを変換されないように拡張子をリネーム
                RenameAllFileExtension dst_path_0, encode_type
                
                Common.CopyFolder dst_path_0, dst_path_1
            Else
                Common.DeleteFolder current_wk_dst_dir_path
                
                '全ファイルを変換されないように拡張子をリネーム
                RenameAllFileExtension current_wk_src_dir_path, encode_type
                
                Common.CopyFolder current_wk_src_dir_path, current_wk_dst_dir_path
            End If
        Else
            Common.DeleteFolder current_wk_dst_dir_path
            Common.CopyFolder current_wk_src_dir_path, current_wk_dst_dir_path
        End If
        
    End If
        
    Common.WriteLog "CopySrcToDstWorkFolder E"
End Sub

'exeに渡すパラメータリスト作成する
Private Function CreateExeParamList(ByRef sub_param As SubParam) As String()
    Common.WriteLog "CreateExeParamList S"

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
        
        src_path_list = Common.DeleteEmptyArray(src_path_list)
        dst_path_list = Common.DeleteEmptyArray(dst_path_list)
    End If
    
    For i = LBound(src_path_list) To UBound(src_path_list)
        ReDim Preserve param_list(i)
        '"srcdirpath" "*.vb" "dstdirpath" "inipath" "ignorefilelistpath" ""
        param_list(i) = _
            DQ & src_path_list(i) & DQ & " " & _
            DQ & main_param.GetInExtension() & DQ & " " & _
            DQ & dst_path_list(i) & DQ & " " & _
            DQ & sub_param.GetIniFilePath() & DQ & " " & _
            DQ & main_param.GetIgnoreFilePath() & DQ & " " & _
            DQ & DQ
    Next i
    
    CreateExeParamList = param_list
    
    Common.WriteLog "CreateExeParamList E"
End Function

'exeを実行する
Private Sub RunExe(ByRef param_list() As String)
    Common.WriteLog "RunExe S"

    Dim i As Integer
    Dim ret As Long
    Dim exe_param As String
    
    For i = LBound(param_list) To UBound(param_list)
        
        exe_param = _
            DQ & main_param.GetExeFilePath() & DQ & " " & _
            param_list(i)
            
        ChDir Common.GetFolderNameFromPath(main_param.GetExeFilePath())
        
        Common.WriteLog exe_param
        
        ret = Common.RunProcessWait(exe_param)
        
        If ret <> 0 Then
            Err.Raise 53, , "Exeの実行に失敗しました(exe ret=" & ret & ")"
        End If
    
    Next i

    Common.WriteLog "RunExe E"
End Sub

'差分があるかチェックする
Private Function IsMatch() As Boolean
    Common.WriteLog "IsMatch S"

    Dim i As Integer
    Dim is_match As Boolean: is_match = True

    'ファイルリストを作成
    Dim src_file_list() As String: src_file_list = Common.CreateFileList(current_wk_src_dir_path, main_param.GetInExtension(), True)
    Dim dst_file_list() As String: dst_file_list = Common.CreateFileList(current_wk_dst_dir_path, main_param.GetInExtension(), True)

    'ファイルを比較
    For i = LBound(src_file_list) To UBound(src_file_list)
    
        If src_file_list(i) = "" Or dst_file_list(i) = "" Then
            Exit For
        End If
        
        is_match = Common.IsMatchTextFiles(src_file_list(i), dst_file_list(i))
        
        If is_match = False Then
            '1つでも差異があれば終了
            IsMatch = is_match
            Common.WriteLog "IsMatch E1"
            Exit Function
        End If
    Next i
    
    IsMatch = is_match
    Common.WriteLog "IsMatch E"
End Function

'作業用サブフォルダを入れ替える
Private Sub SwapWorkSubFolder()
    Common.WriteLog "SwapWorkSubFolder S"
    
    If main_param.IsStepWorkDir() = True Then
        Common.WriteLog "SwapWorkSubFolder E1"
        Exit Sub
    End If
    
    Dim tmp As String: tmp = current_wk_src_dir_path
    current_wk_src_dir_path = current_wk_dst_dir_path
    current_wk_dst_dir_path = tmp
    
    Common.WriteLog "SwapWorkSubFolder E"
End Sub

'作業用フォルダを削除する
Private Sub DeleteWorkFolder(ByVal is_del_wk_dir As Boolean)
    Common.WriteLog "DeleteWorkFolder S"

    If is_del_wk_dir = False Then
        Common.WriteLog "DeleteWorkFolder E1"
        Exit Sub
    End If

    Dim path As String: path = main_param.GetToolWorkDirPath()
    
    If Common.IsExistsFolder(path) = False Then
        Common.WriteLog "DeleteWorkFolder E2"
        Exit Sub
    End If
    
    If Common.ShowYesNoMessageBox("作業用フォルダを削除しますか?") = False Then
        Common.WriteLog "DeleteWorkFolder E3"
        Exit Sub
    End If
    
    Common.DeleteFolder path
    Common.WriteLog "DeleteWorkFolder E"
End Sub

Private Function IsTestMode() As Boolean
    Common.WriteLog "IsTestMode S"
    IsTestMode = False
    If Common.GetFileName(main_param.GetExeFilePath()) = "cs.exe" Then
        Common.WriteLog "[TestMode]"
        IsTestMode = True
    End If
    Common.WriteLog "IsTestMode E"
End Function
