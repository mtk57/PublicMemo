Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String
Private Const VB6_EXT = "bas,frm,cls,ctl"
Private Const VBNET_EXT = "vb"

'パラメータ
Private main_param As MainParam

Private target_files() As String

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    '対象ファイルを検索する
    SearchTargetFile
    
    '対象ファイルを同じフォルダ構造のままコピーする
    CopyTargetFiles
    
    'メインループ
    Dim i As Long
    For i = LBound(target_files) To UBound(target_files)
        Dim targer_path As String: targer_path = target_files(i)
        Common.WriteLog "i=" & i & ":[" & targer_path & "]"
    
        '対象ファイルの関数の先頭と最後にコードを埋め込む
        InsertCode targer_path
    Next i

    Common.WriteLog "Run E"
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

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'対象ファイルを検索する
Private Sub SearchTargetFile()
    Common.WriteLog "SearchTargetFile S"
    
    Dim err_msg As String
    Dim path As String
    Dim i As Long: i = 0
    
    Erase target_files
    
    '対象ファイルを検索する

    '拡張子リスト作成
    Dim ext_list() As String
    If main_param.GetTargetExtension() = "VB6系" Then
        ext_list = Split(VB6_EXT, ",")
    Else
        ext_list = Split(VBNET_EXT, ",")
    End If

    '拡張子でループ
    For i = LBound(ext_list) To UBound(ext_list)
        '拡張子で検索してファイルリスト作成
        Dim temp_list() As String
        temp_list = Common.CreateFileList( _
                        main_param.GetTargetDirPath(), _
                        "*." & ext_list(i), _
                        main_param.IsSubDir() _
                    )
        '結果マージ
        target_files = Common.MergeArray(target_files, temp_list)
    Next i
    
    target_files = Common.DeleteEmptyArray(target_files)
    
    If Common.IsEmptyArray(target_files) = True Then
        err_msg = "対象ファイルが見つかりませんでした"
        Err.Raise 53, , err_msg
    End If
    
    Common.WriteLog "SearchTargetFile E"
End Sub

'対象ファイルを同じフォルダ構造のままコピーする
Private Sub CopyTargetFiles()
    Common.WriteLog "CopyTargetFiles S"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim base_path As String: base_path = Common.GetCommonString(target_files)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim dst_file_path() As String
    Dim i As Integer
    Dim cnt As Integer: cnt = 0
    Dim err_msg As String: err_msg = ""
    
    Common.DeleteFolder main_param.GetDestDirPath()
    
    For i = LBound(target_files) To UBound(target_files)
        Dim src As String: src = target_files(i)
        
        If Common.IsExistsFile(src) = False Then
            err_msg = "ファイルが存在しません" & vbCrLf & _
                      "src=" & src
            Common.WriteLog "[CopyTargetFiles] ★★エラー! err_msg=" & err_msg
            
            If Common.ShowYesNoMessageBox( _
                "[CopyTargetFiles]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
                "err_msg=" & err_msg _
                ) = False Then
                Err.Raise 53, , "[CopyProjectFiles] エラー! (err_msg=" & err_msg & ")"
            End If
            
            GoTo CONTINUE
        End If
        
        If IsIgnoreFile(src) = True Then
            '除外ファイルは除外する
            Common.WriteLog "除外=" & src
            GoTo CONTINUE
        End If
        
        If IsIgnoreKeyword(src) = True Then
            '除外キーワードを含むので除外する
            Common.WriteLog "除外=" & src
            GoTo CONTINUE
        End If
        
        Dim dst As String: dst = main_param.GetDestDirPath() & SEP & dst_base_path & Replace(src, base_path, "")
        Dim path As String: path = Common.GetFolderNameFromPath(dst)
        
        'フォルダが存在しない場合は作成する
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        'ファイルをコピーする
        fso.CopyFile src, dst
        
        ReDim Preserve dst_file_path(cnt)
        dst_file_path(cnt) = dst
        
        cnt = cnt + 1
        
CONTINUE:
        
    Next i
    
    Erase target_files
    target_files = Common.MergeArray(target_files, dst_file_path)
    target_files = Common.DeleteEmptyArray(target_files)
    
    '起点フォルダを移動する
    MoveBaseFolder
    
    Set fso = Nothing
    
    Common.WriteLog "CopyTargetFiles E"
End Sub

'除外ファイルかを返す
Private Function IsIgnoreFile(ByVal path As String) As Boolean
    Common.WriteLog "IsIgnoreFile S"
    
    If main_param.GetIgnoreFiles() = "" Then
        IsIgnoreFile = False
        Common.WriteLog "IsIgnoreFile E1"
        Exit Function
    End If

    '除外ファイルリストを作成
    Dim ignore_files() As String
    ignore_files = Split(main_param.GetIgnoreFiles(), ",")

    If Common.IsEmptyArray(ignore_files) = True Then
        IsIgnoreFile = False
        Common.WriteLog "IsIgnoreFile E2"
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_files) To UBound(ignore_files)
        If Common.GetFileName(path) = ignore_files(i) Then
            IsIgnoreFile = True
            Common.WriteLog "IsIgnoreFile E3"
            Exit Function
        End If
    Next i
    
    IsIgnoreFile = False
    Common.WriteLog "IsIgnoreFile E"
End Function

'除外キーワードを含むかを返す
Private Function IsIgnoreKeyword(ByVal path As String) As Boolean
    Common.WriteLog "IsIgnoreKeyword S"
    
    If main_param.GetIgnoreKeywords() = "" Then
        IsIgnoreKeyword = False
        Common.WriteLog "IsIgnoreKeyword E1"
        Exit Function
    End If

    '除外ファイルリストを作成
    Dim ignore_keywords() As String
    ignore_keywords = Split(main_param.GetIgnoreKeywords(), ",")

    If Common.IsEmptyArray(ignore_keywords) = True Then
        IsIgnoreKeyword = False
        Common.WriteLog "IsIgnoreKeyword E2"
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_keywords) To UBound(ignore_keywords)
        If InStr(Common.GetFileName(path), ignore_keywords(i)) > 0 Then
            IsIgnoreKeyword = True
            Common.WriteLog "IsIgnoreKeyword E3"
            Exit Function
        End If
    Next i
    
    IsIgnoreKeyword = False
    Common.WriteLog "IsIgnoreKeyword E"
End Function

'起点フォルダを移動する
Private Sub MoveBaseFolder()
    Common.WriteLog "MoveBaseFolder S"

    If main_param.GetBaseDir() = "" Then
        Common.WriteLog "MoveBaseFolder E1"
        Exit Sub
    End If
    
    '起点フォルダ名が指定されている場合、コピー先フォルダパスに存在するかチェックする
    Dim base_dir As String: base_dir = ""
    Dim i As Long
    For i = LBound(target_files) To UBound(target_files)
        base_dir = Common.GetFolderPathByKeyword( _
                        Common.GetFolderNameFromPath(target_files(i)), _
                        main_param.GetBaseDir())
        If base_dir <> "" Then
            Exit For
        End If
    Next i
    
    '存在しない場合は何もしない
    If base_dir = "" Then
        Common.WriteLog "MoveBaseFolder E2"
        Exit Sub
    End If
    
    Dim renamed_dir As String: renamed_dir = main_param.GetBaseDir()
    
    '存在する場合は移動する
    If Common.IsExistsFolder(main_param.GetDestDirPath() & SEP & renamed_dir) = True Then
        '移動先に同名フォルダがある場合はユニークなフォルダ名にする
        renamed_dir = Common.GetLastFolderName( _
                            Common.ChangeUniqueDirPath( _
                                main_param.GetDestDirPath() & SEP & renamed_dir))
    End If
    
    Common.MoveFolder base_dir, main_param.GetDestDirPath() & SEP & renamed_dir
    
    '最後にフォルダを削除する
    Dim dust_dir As String: dust_dir = Replace(base_dir, main_param.GetDestDirPath() & SEP, "")
    Dim del_dir_path As String: del_dir_path = main_param.GetDestDirPath() & SEP & Split(dust_dir, SEP)(0)
    
    If Common.IsExistsFolder(del_dir_path) = False Then
        Common.WriteLog "MoveBaseFolder E3"
        Exit Sub
    End If
    
    Common.DeleteFolder del_dir_path
    
    '対象ファイルリストも再作成する
    For i = LBound(target_files) To UBound(target_files)
        Dim new_path As String
        new_path = Replace(target_files(i), base_dir, main_param.GetDestDirPath() & SEP & renamed_dir)
        target_files(i) = new_path
    Next i
    
    Common.WriteLog "MoveBaseFolder E"
End Sub

'対象ファイルの関数の先頭と最後にコードを埋め込む
Private Sub InsertCode(ByVal target_path As String)
    Common.WriteLog "InsertCode S"
    
    Dim contents() As String: contents = GetTargetContents(target_path)
    
    If Common.IsEmptyArray(contents) = True Then
        Common.WriteLog "InsertCode E1"
        Exit Sub
    End If
    
    Const METHOD_START = "(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+.*"

    Dim new_contents() As String
    ReDim new_contents(0)
    Dim i As Long
 
    For i = LBound(contents) To UBound(contents)
        Dim line As String: line = contents(i)
        
        If Common.IsCommentCode(line, Common.GetFileExtension(target_path)) = True Then
            'コメント行なので次の行へ
            GoTo NOT_METHOD
        End If
        
        If Common.IsMatchByRegExp(line, METHOD_START, True) = True Then
        
            If IsExistIgnoreMethodWord(line) = True Then
                '除外ワードを含むので次の行へ
                GoTo NOT_METHOD
            End If
            
            '関数定義の開始行を発見
            i = i + InsertCodeForMethod( _
                        target_path, _
                        i, _
                        contents, _
                        new_contents _
                    )
            
            GoTo CONTINUE
        End If
            
NOT_METHOD:
        '関数定義以外の行
        Common.AppendArray new_contents, line
        
CONTINUE:
    
    Next i
    
    '最後にファイルに出力する
    Common.CreateSJISTextFile new_contents, target_path
    
FINISH:
    Common.WriteLog "InsertCode E"

End Sub

'関数にコードを挿入する
Private Function InsertCodeForMethod( _
    ByVal target_path As String, _
    ByVal start As Long, _
    ByRef contents() As String, _
    ByRef new_contents() As String _
) As Long
    Common.WriteLog "InsertCodeForMethod S"
    
    Const METHOD_END = "End\s(Function|Sub)"
    Const METHOD_EXIT = "Exit\s(Function|Sub)"
    
    Dim i As Long
    Dim line As String: line = contents(start)  '解析中の行データ
    Dim method_name As String: method_name = GetMethodName(line)
    Dim cnt As Long     '解析を進めた行数。ただし開始行および追加行は含まない。
    Dim offset As Long  '関数開始位置のオフセット行数(関数の引数が複数行の場合は2行以上になる)
    Dim seq As Long: seq = 1    '関数途中終了時を区別するための連番

    Common.AppendArray new_contents, line
    cnt = cnt + 1
    
    '関数開始定義の終了行を取得する
    offset = GetMethodStartOffset(target_path, start, contents)
    
    If offset > 0 Then
        For i = 0 To offset - 1
            Common.AppendArray new_contents, contents(start + i + 1)
        Next i
        cnt = cnt + offset - 1
    End If
    Common.AppendArray new_contents, GetMethodStartLine(method_name)
    
    For i = start + offset + 1 To UBound(contents)
        line = contents(i)
        
        If Common.IsCommentCode(line, Common.GetFileExtension(target_path)) = True Then
            'コメント行なので次の行へ
            GoTo METHOD_BODY
        End If
        
        If Common.IsMatchByRegExp(line, METHOD_EXIT, True) = True Then
            '関数の途中終了行を発見
            
            Common.AppendArray new_contents, GetMethodExitLine(method_name, seq)
            Common.AppendArray new_contents, line
            cnt = cnt + 1
            
            seq = seq + 1
            
            GoTo CONTINUE
        End If
        
        If Common.IsMatchByRegExp(line, METHOD_END, True) = True Then
            '関数定義の終了行を発見
            
            Common.AppendArray new_contents, GetMethodEndLine(method_name)
            Common.AppendArray new_contents, line
            cnt = cnt + 1
            
            GoTo FINISH
        End If

METHOD_BODY:
        '関数定義の本体
        Common.AppendArray new_contents, line
        cnt = cnt + 1
        
CONTINUE:
    Next i

FINISH:
    InsertCodeForMethod = cnt
    Common.WriteLog "InsertCodeForMethod E"
End Function

'関数開始定義の終了行を取得する
'
' <考え方>
'  1." "でSplit
'  2."Sub"があればSubモードON, "Function"があればFunctionモードON   ※このメソッドに渡す前に正規表現でヒットした文字列を渡しているので無いことは有り得ない。
'  3.行ループ開始
'  4.  列ループ開始
'  4-1.  列終端に"_"があれば行ループ続行。なければ処理終了する。行ループした回数を終了行とする。
'  4-2.  "("があれば括弧カウンタ++、")"があれば括弧カウンタ--する。
'  4-3.  括弧カウンタが0になった && SubモードならSubの終わりと判断し処理終了。行ループした回数を終了行とする。
'  4-4.  括弧カウンタが0になった && FunctionモードならFunctionの引数の終わりと判断するが、戻り値がある可能性があるので戻り値モードONして処理続行する。
'  4-5.  戻り値モードON && " As "があれば戻り値がある。列終端に"_"がなければ処理終了。行ループした回数を終了行とする。
'
'  - 複数行の場合、"_"以降の列にコメントや" "は付けられない。
'  - 複数行の場合、コメントは一切付けられない。
'  - 引数も戻り値も無いFunctionを作成可能
'  - 戻り値が配列の場合"()"で終わるので、Functionモード時の括弧カウンタには注意が必要。
'  - 上記の<考え方>は正常ケースのみ(つまり正常にビルドできるコード)。
'
Private Function GetMethodStartOffset( _
    ByVal target_path As String, _
    ByVal start As Long, _
    ByRef contents() As String _
) As Long
    Common.WriteLog "GetMethodStartOffset S"
    
    Dim offset As Long: offset = 0
    Dim r As Long       'Row
    Dim c As Long       'Column
    Dim line As String
    Dim ch As String    'Character
    Dim kc As Long: kc = -1      '括弧カウンタ(-1:1つ目の"("が未発見、-2:戻り値より前の括弧は全ての対応確認済(戻り値モードON))
    Dim mode As String: mode = GetMethodType(contents(start))   '関数モード("Sub" or "Function")
    
    For r = start To UBound(contents)
        line = contents(r)
        
        If Right(line, 1) <> "_" Then
            '複数行では無いので処理終了
            GoTo FINISH
        End If
        
        For c = 1 To Len(line)
            ch = Mid(line, c, 1)    '1文字取得
        
            '戻り値より前の括弧をチェック
            If ch = "(" Then
                If kc = -1 Then
                    kc = 0
                End If
                If kc >= 0 Then
                    kc = kc + 1
                End If
            ElseIf ch = ")" Then
                If kc >= 1 Then
                    kc = kc - 1
                End If
            End If
            
            If kc = 0 Then
                '戻り値より前の括弧は全ての対応確認済
                If mode = "Sub" Then
                    'Subの場合は処理終了
                    GoTo FINISH
                End If
                kc = -2 '戻り値モードON
            End If
            
        Next c

        offset = offset + 1
        
    Next r

    Err.Raise 53, , "関数開始定義の終了行が見つかりません (" & target_path & ")"
    
FINISH:
    GetMethodStartOffset = offset
    Common.WriteLog "GetMethodStartOffset E"
End Function

Private Function GetMethodType(ByVal line As String) As String
    Common.WriteLog "GetMethodType S"
    
    Dim i As Long
    Dim start_rows() As String
    start_rows = Split(line, " ")
    GetMethodType = ""
    
    For i = 0 To UBound(start_rows)
        If start_rows(i) = "Sub" Or start_rows(i) = "Function" Then
            GetMethodType = start_rows(i)
            Exit For
        End If
    Next i
    
    If GetMethodType = "" Then
        Err.Raise 53, , "関数開始定義が不正です (line=" & line & ")"
    End If
        
    Common.WriteLog "GetMethodType E"
End Function

'関数開始直後に挿入するコードを作成する
Private Function GetMethodStartLine(ByVal method_name As String) As String
    Common.WriteLog "GetMethodStartLine S"
    GetMethodStartLine = Replace(main_param.GetInsertWord(), "＠", method_name & " START")
    Common.WriteLog "GetMethodStartLine E"
End Function

'関数終了直前に挿入するコードを作成する
Private Function GetMethodEndLine(ByVal method_name As String) As String
    Common.WriteLog "GetMethodEndLine S"
    GetMethodEndLine = Replace(main_param.GetInsertWord(), "＠", method_name & " END")
    Common.WriteLog "GetMethodEndLine E"
End Function

'関数途中直前に挿入するコードを作成する
Private Function GetMethodExitLine(ByVal method_name As String, ByVal seq As Long) As String
    Common.WriteLog "GetMethodExitLine S"
    Common.WriteLog "seq=" & seq
    GetMethodExitLine = Replace(main_param.GetInsertWord(), "＠", method_name & " END " & seq)
    Common.WriteLog "GetMethodExitLine E"
End Function

'関数名を返す
'Function,Subのすぐ後ろに関数名がある場合のみ想定。複数行は未対応。
Private Function GetMethodName(ByVal line As String) As String
    Common.WriteLog "GetMethodName S"
    
    Const METHOD = "\s*(Function|Sub)\s+.*"
    
    Dim list() As String
    list = Common.GetMatchByRegExp(line, METHOD, True)
    
    list = Common.DeleteEmptyArray(list)
    list = Split(list(0), " ")
    
    Dim method_name As String
    
    Dim i As Long
    For i = 0 To UBound(list)
        If list(i) = "Sub" Or list(i) = "Function" Then
            method_name = list(i + 1)
            
            '括弧以降を除去
            GetMethodName = Replace( _
                                method_name, _
                                Common.GetStringByKeyword(method_name, "("), _
                                "" _
                            )
            
            Common.WriteLog "GetMethodName E1"
            Exit Function
        End If
    Next
    
    Err.Raise 53, , "関数名が見つかりません (" & line & ")"
    
    Common.WriteLog "GetMethodName E"
End Function

'関数定義に不要はワードがあるか返す
Private Function IsExistIgnoreMethodWord(ByVal line As String) As Boolean
    Common.WriteLog "IsExistIgnoreMethodWord S"

    Const IGNORE_WORDS = "Declare,PtrSafe,Lib,Alias"

    Dim list() As String
    Dim ignores() As String
    list = Split(line, " ")
    ignores = Split(IGNORE_WORDS, ",")
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(list)
        For j = 0 To UBound(ignores)
            If list(i) = ignores(j) Then
                '除外ワード発見
                IsExistIgnoreMethodWord = True
                Common.WriteLog "IsExistIgnoreMethodWord E1"
                Exit Function
            End If
        Next j
    Next i
    
    IsExistIgnoreMethodWord = False
    
    Common.WriteLog "IsExistIgnoreMethodWord E"
End Function

'対象ファイルを読み込んで内容を配列で返す
Private Function GetTargetContents(ByVal path As String) As String()
    Common.WriteLog "GetTargetContents S"
    
    Dim raw_contents As String
    Dim contents() As String
    
    'ファイルを開いて、全行を配列に格納する
    If Common.IsSJIS(path) = True Then
        raw_contents = Common.ReadTextFileBySJIS(path)
    ElseIf Common.IsUTF8(path) = True Then
        raw_contents = Common.ReadTextFileByUTF8(path)
    Else
        Dim err_msg As String: err_msg = "未サポートのエンコードです" & vbCrLf & _
                  "path=" & path
        Common.WriteLog "[GetTargetContents] ★★エラー! err_msg=" & err_msg
        
        If Common.ShowYesNoMessageBox( _
            "[GetTargetContents]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetTargetContents] エラー! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog "GetTargetContents E1"
        GetTargetContents = contents
        Exit Function
    End If
    
    contents = Split(raw_contents, vbCrLf)
    
    GetTargetContents = contents

    Common.WriteLog "GetTargetContents E"
End Function


