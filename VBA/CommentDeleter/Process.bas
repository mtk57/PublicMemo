Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String
'Private Const VB6_EXT = "bas,frm,cls,ctl"
'Private Const VBNET_EXT = "vb"
Private Const PLSQL_EXT = "sql"
'Private Const REPLACE_METHOD = "＠"
'Private Const REPLACE_FILENAME = "＊"

'Private Const METHOD = "\s*(Function|Sub)\s+.*"
'Private Const METHOD_START = "(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+.*\("
'Private Const METHOD_END = "End\s(Function|Sub)$"
'Private Const METHOD_EXIT = "Exit\s(Function|Sub)$"
'Private Const METHOD_APP_END = "^(\t|\s)*\bEnd$"
'Private Const METHOD_RET = "^[ \t]*(Return|Throw) *"

'Private Const IGNORE_WORDS = "Declare,PtrSafe,Lib,Alias"

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
    
        '対象ファイルのコメントを削除する
        DeleteComment targer_path
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
    If main_param.GetTargetExtension() = "PL/SQL(sql)" Then
        ext_list = Split(PLSQL_EXT, ",")
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
        
        If main_param.IsIgnoreFile(src) = True Then
            '除外ファイルは除外する
            Common.WriteLog "除外ファイル=" & src
            GoTo CONTINUE
        End If
        
        If main_param.IsIgnoreKeyword(src) = True Then
            '除外キーワードを含むので除外する
            Common.WriteLog "除外ファイル=" & src
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

'対象ファイルのコメントを削除する
Private Sub DeleteComment(ByVal target_path As String)
    Common.WriteLog "InsertCode S"
    
    Dim contents() As String: contents = GetTargetContents(target_path)
    
    If Common.IsEmptyArrayLong(contents) = True Then
        Common.WriteLog "DeleteComment E1"
        Exit Sub
    End If
    
    Dim new_contents() As String
    
    If main_param.GetTargetExtension() = "PL/SQL(sql)" Then
        'PL/SQLコメントの削除
        new_contents = RemovePLSQLComments(contents)
    End If
    
    '最後にファイルに出力する
    Common.CreateSJISTextFile new_contents, target_path
    
FINISH:
    Common.WriteLog "DeleteComment E"

End Sub

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

Private Function RemovePLSQLComments(sqlLines() As String) As String()
    Common.WriteLog "RemovePLSQLComments S"

    Dim result() As String
    Dim resultCount As Long
    Dim inString As Boolean
    Dim inBlockComment As Boolean
    Dim i As Long, j As Long
    Dim currentChar As String
    Dim nextChar As String
    Dim currentLine As String
    Dim processedLine As String
    
    ' 結果配列の初期化
    ReDim result(LBound(sqlLines) To UBound(sqlLines))
    resultCount = 0
    
    inString = False
    inBlockComment = False
    
    ' 各行を処理
    For i = LBound(sqlLines) To UBound(sqlLines)
        currentLine = sqlLines(i)
        processedLine = ""
        
        j = 1
        Do While j <= Len(currentLine)
            currentChar = Mid(currentLine, j, 1)
            
            If j < Len(currentLine) Then
                nextChar = Mid(currentLine, j + 1, 1)
            Else
                nextChar = ""
            End If
            
            ' 文字列リテラルの処理
            If currentChar = "'" And Not inBlockComment Then
                ' エスケープされたクォート('')の処理
                If nextChar = "'" Then
                    processedLine = processedLine & "''"
                    j = j + 2
                    GoTo ContinueLoop
                Else
                    inString = Not inString
                    processedLine = processedLine & currentChar
                End If
            
            ' ブロックコメント開始の処理
            ElseIf currentChar = "/" And nextChar = "*" And Not inString And Not inBlockComment Then
                inBlockComment = True
                j = j + 1 ' 次の文字(*)をスキップ
            
            ' ブロックコメント終了の処理
            ElseIf currentChar = "*" And nextChar = "/" And inBlockComment Then
                inBlockComment = False
                j = j + 1 ' 次の文字(/)をスキップ
            
            ' 行コメントの処理
            ElseIf currentChar = "-" And nextChar = "-" And Not inString And Not inBlockComment Then
                ' 行コメントの場合、この行の残りは無視
                Exit Do
            
            ' コメント内でなければ文字を結果に追加
            ElseIf Not inBlockComment Then
                processedLine = processedLine & currentChar
            End If
            
            j = j + 1
ContinueLoop:
        Loop
        
        ' 処理された行を結果配列に追加
        result(resultCount) = processedLine
        resultCount = resultCount + 1
    Next i
    
    ' 結果配列のサイズを実際の結果数に調整
    ReDim Preserve result(LBound(sqlLines) To LBound(sqlLines) + resultCount - 1)
    
    RemovePLSQLComments = result
    
    Common.WriteLog "RemovePLSQLComments E"
End Function
