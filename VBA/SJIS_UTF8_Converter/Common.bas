Attribute VB_Name = "Common"
Option Explicit

Public Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Public Declare PtrSafe Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

'ログファイル番号
Private logfile_num As Integer
Private is_log_opened As Boolean

'-------------------------------------------------------------
'指定されたファイルをSJIS→UTF8(BOMあり)変換する
' path : IN : ファイルパス(絶対パス)
' is_backup : IN : True/False (True=バックアップする)
'                  →末尾に".bak_現在日時"を付与
'-------------------------------------------------------------
Public Sub SJIStoUTF8(ByVal path As String, ByVal is_backup As Boolean)
    Dim in_str As String
    Dim buf As String
    Dim i As Integer
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS形式のテキストファイルを読み込み
    in_str = ""
    Open path For Input As #filenum
        'テキストをすべて取得する
        Do Until EOF(filenum)
            Line Input #filenum, buf
            in_str = in_str & buf & vbLf
        Loop
    Close #filenum
        
    'Shift-JIS以外のファイルを読み込んでしまった場合は終了
    For i = 1 To Len(in_str)
        If Asc(Mid(in_str, i, 1)) = -7295 Then Exit Sub
    Next
    
    'バックアップ
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    'UTF-8（BOM付き）でテキストファイルへ出力
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText in_str, 0
        .SaveToFile path, 2
        .Close
    End With
    
End Sub

'-------------------------------------------------------------
'指定されたファイルをUTF8(BOMあり/なし) → SJIS変換する
' path : IN : ファイルパス(絶対パス)
' is_backup : IN : True/False (True=バックアップする)
'                  →末尾に".bak_現在日時"を付与
'-------------------------------------------------------------
Public Sub UTF8toSJIS(ByVal path As String, ByVal is_backup As Boolean)
    Dim in_str As String
    Dim out_str() As String
    Dim i As Integer
    
    'UTF-8もしくはUTF-8（BOM付き）のテキストファイルを読み込み
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8もしくはUTF-8（BOM付き）以外を読み込んでしまった場合は終了
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                Exit Sub
            End If
        End If
    Next
    
    '改行毎にデータを分ける
    out_str = Split(in_str, vbLf)
    
    'バックアップ
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS形式でテキストファイルへ出力
    Open path For Output As #filenum
        For i = 0 To UBound(out_str)
            Print #filenum, out_str(i)
        Next
    Close #filenum

End Sub

'-------------------------------------------------------------
'ファイルがUTF8(BOMあり)かを判定する
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=UTF8(BOMあり), False=左記以外)
'-------------------------------------------------------------
Public Function IsUTF8(ByVal path As String) As Boolean
    If Common.IsExistsFile(path) = False Then
        Err.Raise 53, , "指定されたファイルが存在しません (" & path & ")"
    End If

    Dim bytedata() As Byte: bytedata = ReadBinary(path, 3)
    Dim length As Integer: length = UBound(bytedata) + 1
    
    If length < 3 Then
        IsUTF8 = False
        Exit Function
    End If
    
    If bytedata(0) = &HEF And bytedata(1) = &HBB And bytedata(2) = &HBF Then
        IsUTF8 = True
    Else
        IsUTF8 = False
    End If
    
End Function

'-------------------------------------------------------------
'ファイルをバイナリとして指定サイズ読み込む
' path : IN : ファイルパス(絶対パス)
' readsize : IN : 読み込むサイズ
' Ret : 読み込んだバイナリ配列
'-------------------------------------------------------------
Public Function ReadBinary(ByVal path As String, ByVal readsize As Integer) As Byte()
    Dim readdata() As Byte
    
    If readsize <= 0 Then
        ReadBinary = readdata()
        Exit Function
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    Open path For Binary Access Read As #filenum
    
    ReDim readdata(readsize - 1)
    
    Get #filenum, , readdata
    
    Close #filenum
    
    ReadBinary = readdata
End Function

'-------------------------------------------------------------
'指定フォルダ配下に指定拡張子のファイルが存在するか
' path : IN : フォルダパス(絶対パス)
' ext : IN : 拡張子(Ex. ".vb")
' Ret : True/False (True=存在する, False=存在しない)
'-------------------------------------------------------------
Public Function IsExistsExtensionFile(ByVal path As String, ByVal ext As String) As Boolean
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim File As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    For Each subfolder In folder.subfolders
        If IsExistsExtensionFile(subfolder.path, ext) Then
            Set fso = Nothing
            Set folder = Nothing
            
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next subfolder
    
    For Each File In folder.Files
        If Right(File.Name, Len(ext)) = ext Then
            Set fso = Nothing
            Set folder = Nothing
        
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next File
    
    Set fso = Nothing
    Set folder = Nothing

    IsExistsExtensionFile = False
End Function

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
'配列の空行を削除する
' in_array : IN : 文字列配列
' Ret : 空行を削除した配列
'-------------------------------------------------------------
Public Function DeleteEmptyArray(ByRef in_array() As String) As String()
    Dim ret_array() As String
    Dim i, cnt As Long
    Dim row As String
    
    ReDim ret_array(UBound(in_array))
    
    For i = LBound(in_array) To UBound(in_array)
        row = in_array(i)
        If Not IsEmpty(row) Then
            If row <> "" Then
                ret_array(cnt) = row
                cnt = cnt + 1
            End If
        End If
    Next
    
    ReDim Preserve ret_array(cnt - 1)
    
    DeleteEmptyArray = ret_array
End Function

'-------------------------------------------------------------
'ファイルリストを作成する
' path : IN : フォルダパス(絶対パス)
' ext : IN : 拡張子(Ex."*.vb")
' is_subdir : IN : サブフォルダ含むか (True=含む)
' Ret : ファイルリスト
'-------------------------------------------------------------
Public Function CreateFileList(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean) As String()
    Dim list() As String: list = CreateFileListMain(path, ext, is_subdir)
    CreateFileList = DeleteEmptyArray(list)
End Function

Private Function CreateFileListMain(ByVal path As String, ByVal ext As String, ByVal is_subdir As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filelist() As String
    Dim cnt As Integer

    Dim File As String, f As Object
    File = Dir(path & "\" & ext)
    
    If File <> "" Then
        If Common.IsEmptyArray(filelist) = True Then
            cnt = 0
        Else
            cnt = UBound(filelist) + 1
        End If
        
        ReDim Preserve filelist(cnt)
        filelist(cnt) = path & "\" & File
    End If
    
    Do While File <> ""
        File = Dir()
        If File <> "" Then
            cnt = UBound(filelist) + 1
            ReDim Preserve filelist(cnt)
            filelist(cnt) = path & "\" & File
        End If
    Loop
    
    If is_subdir = False Then
        Set fso = Nothing
        CreateFileListMain = filelist
        Exit Function
    End If
    
    Dim filelist_sub() As String
    Dim filelist_merge() As String
    
    For Each f In fso.GetFolder(path).subfolders
        filelist_sub = CreateFileListMain(f.path, ext, is_subdir)
        filelist = Common.MergeArray(filelist_sub, filelist)
    Next f
    
    Set fso = Nothing
    CreateFileListMain = filelist
End Function

'-------------------------------------------------------------
'2つの配列を結合して返す
' array1 : IN : 配列1
' array2 : IN : 配列2
' Ret : 結合した配列
'-------------------------------------------------------------
Public Function MergeArray(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim merged As Variant
    merged = Split(Join(array1, vbCrLf) & vbCrLf & Join(array2, vbCrLf), vbCrLf)
    MergeArray = merged
End Function

'-------------------------------------------------------------
'2つのテキストファイルを比較して一致しているかを返す
' file1 : IN : ファイル1パス(絶対パス)
' file2 : IN : ファイル2パス(絶対パス)
' Ret : 比較結果 : True/False (True=一致)
'-------------------------------------------------------------
Public Function IsMatchTextFiles(ByVal file1 As String, ByVal file2 As String) As Boolean
    Dim filesize1 As Long: filesize1 = FileLen(file1)
    Dim filesize2 As Long: filesize2 = FileLen(file2)
    
    'まずファイルサイズでチェック
    If filesize1 = 0 And filesize2 = 0 Then
        'どちらも0byteなので一致
        IsMatchTextFiles = True
        Exit Function
    ElseIf filesize1 <> filesize2 Then
        'ファイルサイズが異なるので不一致
        IsMatchTextFiles = False
        Exit Function
    ElseIf filesize1 = 0 Or filesize2 = 0 Then
        'どちらかが0byteなので不一致
        IsMatchTextFiles = False
        Exit Function
    End If

    Dim fso1, fso2 As Object
    Dim ts1, ts2 As Object
    
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    
    Const READ_ONLY = 1
    Set ts1 = fso1.OpenTextFile(file1, READ_ONLY)
    Set ts2 = fso2.OpenTextFile(file2, READ_ONLY)
    
    Dim contents1 As String: contents1 = ts1.ReadAll
    Dim contents2 As String: contents2 = ts2.ReadAll
    
    ts1.Close
    ts2.Close
    Set ts1 = Nothing
    Set ts2 = Nothing
    Set fso1 = Nothing
    Set fso2 = Nothing
    
    IsMatchTextFiles = (contents1 = contents2)
End Function

'-------------------------------------------------------------
'文字列の配列の末尾に文字列を追加する
' ary : IN/OUT : 文字列の配列
' value : IN : 追加する文字列
'-------------------------------------------------------------
Public Sub AppendArray(ByRef ary() As String, ByVal value As String)
    Dim cnt As Integer: cnt = UBound(ary) + 1
    ReDim Preserve ary(cnt)
    ary(cnt) = value
End Sub

'-------------------------------------------------------------
'フォルダパスを列挙する。（サブフォルダ含む）
' 注意：pathは戻り値には含まない
' path : IN : フォルダパス（絶対パス）
' Ret : フォルダパスリスト
'-------------------------------------------------------------
Public Function GetFolderPathList(ByVal path As String) As String()
    Dim fso As Object
    Dim top_dir As Object
    Dim sub_dir As Object
    Dim path_list() As String
    Dim dir_cnt As Long
    Dim i, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set top_dir = fso.GetFolder(path)

    dir_cnt = top_dir.subfolders.count
    If dir_cnt > 0 Then
        ReDim path_list(dir_cnt - 1)
        i = 0
        For Each sub_dir In top_dir.subfolders
            path_list(i) = sub_dir.path
            i = i + 1
            
            Dim sub_path_list() As String
            sub_path_list = GetFolderPathList(sub_dir.path)
            
            'サブフォルダ内のパスを配列に追加する
            If sub_path_list(0) <> "" Then
                Dim cnt As Integer: cnt = UBound(path_list) + UBound(sub_path_list) + 1
                ReDim Preserve path_list(cnt)
                For j = LBound(sub_path_list) To UBound(sub_path_list)
                    path_list(i) = sub_path_list(j)
                    i = i + 1
                Next j
            End If
        Next sub_dir
        
        GetFolderPathList = path_list
    Else
        Dim ret_empty(0) As String
        GetFolderPathList = ret_empty
    End If
    
    Set sub_dir = Nothing
    Set top_dir = Nothing
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'フォルダをコピーする(サブフォルダ含む)
' src_path : IN : コピー元フォルダパス(絶対パス)
' dst_path : IN : コピー先フォルダパス(絶対パス)
'-------------------------------------------------------------
Public Sub CopyFolder(ByVal src_path As String, dest_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'コピー元のフォルダが存在しない場合、エラーを発生させる
    If Not fso.FolderExists(src_path) Then
        Err.Raise 53, , "指定されたフォルダが存在しません"
    End If
    
    'コピー先のフォルダが存在しない場合、作成する
    If Not fso.FolderExists(dest_path) Then
        fso.CreateFolder dest_path
    End If
    
    'コピー元のフォルダ内のファイルをコピーする
    Dim File As Object
    For Each File In fso.GetFolder(src_path).Files
        fso.CopyFile File.path, fso.BuildPath(dest_path, File.Name), True
    Next
    
    'コピー元のフォルダ内のサブフォルダをコピーする
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(src_path).subfolders
        CopyFolder subfolder.path, fso.BuildPath(dest_path, subfolder.Name)
    Next
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'Yes/Noメッセージボックスを表示する
' msg : IN : メッセージ
' Ret : True/False (True=Yes, False=No)
'-------------------------------------------------------------
Public Function ShowYesNoMessageBox(ByVal msg As String) As Boolean
    Dim result As Integer: result = MsgBox(msg, vbYesNo, "Confirm")
    
    If result = vbYes Then
        ShowYesNoMessageBox = True
    Else
        ShowYesNoMessageBox = False
    End If
End Function

'-------------------------------------------------------------
'外部アプリケーションを実行し、終了するまで待機する
' exe_path : IN : 外部アプリケーション(exe)の絶対パス
'                 exeに渡すパラメータがある場合も一緒に書くこと
' Ret : プロセスの戻り値
'-------------------------------------------------------------
Public Function RunProcessWait(ByVal exe_path As String) As Long

  Dim wsh As Object
  Set wsh = CreateObject("Wscript.Shell")
  
  Const NOT_DISP = 0
  Const DISP = 1
  Const WAIT = True
  Const NO_WAIT = False
  
  Dim Process As Object
  Set Process = wsh.Exec(exe_path)

  'プロセス完了時に通知を受け取る
  Do While Process.Status = 0
    DoEvents
  Loop

  'プロセスの戻り値を取得する
  RunProcessWait = Process.ExitCode

  Set Process = Nothing
  Set wsh = Nothing
End Function

'-------------------------------------------------------------
'パス文字列の末尾の\を除去して返す
' path : IN : パス文字列
' Ret : パス文字列
'-------------------------------------------------------------
Public Function RemoveTrailingBackslash(ByVal path As String) As String
    If Right(path, 1) = "\" Then
        path = Left(path, Len(path) - 1)
    End If
    RemoveTrailingBackslash = path
End Function

'-------------------------------------------------------------
'ファイルの内容を指定されたシートに出力する
' file_path : IN : ファイルパス (絶対パス)
' sheet_name : IN : シート名
' is_sjis : IN :検索ファイルのエンコード指定。True/False (True=Shift-JIS, False=UTF-16)  TODO:いずれ自動判別したいが。。。
'-------------------------------------------------------------
Public Sub OutputTextFileToSheet(ByVal file_path As String, ByVal sheet_name As String, ByVal is_sjis As Boolean)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'ファイルを開く
    Dim file_format As Integer
    Const FORMAT_ASCII = 0
    Const FORMAT_UNICODE = -1
    
    If is_sjis = True Then
        file_format = FORMAT_ASCII
    Else
        file_format = FORMAT_UNICODE
    End If
    
    Dim File As Object
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    
    Set File = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, file_format)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    'ファイルの内容をシートに出力
    Dim row As Integer: row = 1
    
    Do While Not File.AtEndOfStream
        ws.Cells(row, 1).value = File.ReadLine
        row = row + 1
    Loop
    
    File.Close
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'SJISでテキストファイルを作成する
' contents : IN : 内容
' path : IN : ファイルパス (絶対パス)
'-------------------------------------------------------------
Public Sub CreateSJISTextFile(ByRef contents() As String, ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim txt As Object
    Dim i As Integer
    
    Dim IS_OVERWRITE As Boolean: IS_OVERWRITE = True
    Dim IS_UNICODE As Boolean: IS_UNICODE = False
    
    Set txt = fso.CreateTextFile(path, IS_OVERWRITE, IS_UNICODE)
    
    For i = LBound(contents) To UBound(contents)
        txt.WriteLine contents(i)
    Next i
    
    txt.Close
    Set fso = Nothing
End Sub


'-------------------------------------------------------------
'サブフォルダをまとめて作成する
' path : IN : フォルダパス (絶対パス)
'-------------------------------------------------------------
Public Sub CreateFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folders() As String
    folders = Split(path, Application.PathSeparator)
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(folders) To UBound(folders)
        ary = folders
        ReDim Preserve ary(i)
        If Not fso.FolderExists(Join(ary, Application.PathSeparator)) Then
            Call fso.CreateFolder(Join(ary, Application.PathSeparator))
        End If
    Next
  
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'フォルダを削除する
' path : IN : フォルダパス (絶対パス)
'-------------------------------------------------------------
Public Sub DeleteFolder(ByVal path As String)
    If IsExistsFolder(path) = False Then
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.DeleteFolder path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'文字列配列の共通文字列を返す
' list : IN : 文字列配列
' Ret : 共通文字列
'       Ex. list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'           Ret = "hoge"
'-------------------------------------------------------------
Function GetCommonString(ByRef list() As String) As String
    Dim common_string As String
    Dim i, j As Integer
    Dim flag As Boolean
    
    '最初の文字列を共通文字列の初期値とする
    common_string = list(0)
    
    '各文字列を比較する
    For i = 1 To UBound(list)
        flag = False
        '共通部分を取得する
        For j = 1 To Len(common_string)
            If Mid(common_string, j, 1) <> Mid(list(i), j, 1) Then
                common_string = Left(common_string, j - 1)
                flag = True
                Exit For
            End If
        Next j
    Next i
    
    '結果を出力する
    GetCommonString = common_string
End Function

'-------------------------------------------------------------
'絶対ファイルパスの親フォルダパスを取得する
' path : IN : ファイルパス (絶対パス)
' Ret : 親フォルダパス (絶対パス)
'       Ex. path = "C:\tmp\abc.txt"
'           Ret = "C:\tmp"
'-------------------------------------------------------------
Public Function GetFolderNameFromPath(ByVal path As String) As String
    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = Left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'-------------------------------------------------------------
'相対パスを絶対パスに変換する
' base_path : IN : 基準となるフォルダパス(絶対パス)
' ref_path : IN : ファイルパス（相対パス)
' Ret : 絶対パス
'       Ex. base_path = "C:\tmp\abc"
'           ref_path = "..\cdf\xyz.txt"
'           Ret = "C:\tmp\cdf\xyz.txt"
'-------------------------------------------------------------
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイルの存在チェックを行う
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistsFile(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        IsExistsFile = True
    Else
        IsExistsFile = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'フォルダの存在チェックを行う
' path : IN : フォルダパス(絶対パス)
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistsFolder(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイル名から拡張子を返す
' filename : IN : ファイル名
' Ret : ファイル名の拡張子
'        Ex. "abc.txt"の場合、"txt"が返る
'            "."が含まれていない場合は""が返る
'-------------------------------------------------------------
Public Function GetFileExtension(ByVal filename As String) As String
    Dim dot_pos As Integer
    
    ' "."の位置を取得
    dot_pos = InStrRev(filename, ".")
    
    ' 拡張子を取得
    If dot_pos > 0 Then
        GetFileExtension = Right(filename, Len(filename) - dot_pos)
    Else
        GetFileExtension = ""
    End If
End Function

'-------------------------------------------------------------
'指定フォルダ配下を指定ファイル名で検索してその内容を返す
' target_folder : IN :検索フォルダパス(絶対パス)
' target_file : IN :検索ファイル名
' is_sjis : IN :検索ファイルのエンコード指定。True/False (True=Shift-JIS, False=UTF-8)  TODO:いずれ自動判別したいが。。。
' Ret : 読み込んだファイルの内容
'       配列の末尾には検索ファイルの絶対パスを格納する
'-------------------------------------------------------------
Public Function SearchAndReadFiles(ByVal target_folder As String, ByVal target_file As String, ByVal is_sjis As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim File As Object
    For Each File In folder.Files
        If fso.FileExists(File.path) And fso.GetFileName(File.path) Like target_file Then
            '検索対象のファイルを読み込む
            Dim contents As String
            
            If is_sjis = True Then
                'S-JIS
                contents = ReadTextFileBySJIS(File.path)
            Else
                'UTF-8
                contents = ReadTextFileByUTF8(File.path)
            End If
            
            'ファイルの内容を配列に格納する
            Dim lines() As String: lines = Split(contents, vbCrLf)
            
            '末尾にファイルパスを追加する
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = File.path
            SearchAndReadFiles = lines
            Set fso = Nothing
            Exit Function
        End If
    Next File
    
    'サブフォルダも検索する
    Dim subfolder As Object
    For Each subfolder In folder.subfolders
        Dim result() As String
        result = SearchAndReadFiles(subfolder.path, target_file, is_sjis)
        If UBound(result) >= 1 Then
            'サブフォルダから結果が返ってきた場合は、その結果を返す
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '検索対象のファイルが見つからなかった場合は、空の配列を返す
    Dim ret_empty(0) As String
    SearchAndReadFiles = ret_empty
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'SJIS形式のテキストファイルを読み込む
' file_path : IN : ファイルパス (絶対パス)
' Ret : 読み込んだ内容
'-------------------------------------------------------------
Public Function ReadTextFileBySJIS(ByVal file_path) As String
    'TODO:引数チェック
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const FORMAT_ASCII = 0
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    Dim contents As String
    
    Dim File As Object
    Set File = fso.OpenTextFile(file_path, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    contents = File.ReadAll
    
    File.Close
    Set File = Nothing
    Set fso = Nothing
    
    ReadTextFileBySJIS = contents
End Function

'-------------------------------------------------------------
'UTF-8形式のテキストファイルを読み込む
' file_path : IN : ファイルパス (絶対パス)
' Ret : 読み込んだ内容
'-------------------------------------------------------------
Public Function ReadTextFileByUTF8(ByVal file_path) As String
    'TODO:引数チェック
    
    Dim contents As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file_path
        contents = .ReadText
        .Close
    End With
    
    ReadTextFileByUTF8 = contents
End Function

'-------------------------------------------------------------
'配列が空かをチェックする
' arg : IN : 配列
' Ret : True/False (True=空)
'-------------------------------------------------------------
Public Function IsEmptyArray(arg As Variant) As Boolean
    On Error Resume Next
    IsEmptyArray = Not (UBound(arg) > 0)
    IsEmptyArray = CBool(Err.Number <> 0)
End Function

'-------------------------------------------------------------
'現在日時を文字列で返す
' Ret :Ex."20230326123456"
'-------------------------------------------------------------
Public Function GetNowTimeString() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString = str_date & str_time
End Function

'-------------------------------------------------------------
'シートの存在チェック
' sheet_name : IN : シート名
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistSheet(ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name = sheet_name Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'-------------------------------------------------------------
'シートを追加する
' sheet_name : IN : シート名
'-------------------------------------------------------------
Public Sub AddSheet(ByVal sheet_name As String)
    If IsExistSheet(sheet_name) = True Then
        Application.DisplayAlerts = False
        Sheets(sheet_name).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheet_name
End Sub

