Attribute VB_Name = "Common"
Option Explicit

'サブフォルダをまとめて作成する
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

'文字列配列の共通文字列を返す
'Ex.
'  list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'  Ret = "hoge"
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
        
        '共通部分がない場合はループを終了する
        If flag = False Then
            Exit For
        End If
    Next i
    
    '結果を出力する
    GetCommonString = common_string
End Function

'絶対ファイルパスの親フォルダパスを取得する
'Ex.  path=C:\tmp\abc.txt
'     Ret=C:\tmp
Public Function GetFolderNameFromPath(ByVal path As String) As String
    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = Left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'相対パスを絶対パスに変換する
'Ex.  base_path=C:\tmp\abc
'     ref_path=..\cdf\xyz.txt
'     Ret=C:\tmp\cdf\xyz.txt
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'フォルダの存在チェックを行う
'pathは絶対パスとする
Public Function IsExistsFolder(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'ファイル名から拡張子を返す
'Ex. "abc.txt"の場合、"txt"が返る
'"."が含まれていない場合は""が返る
Public Function GetFileExtension(filename As String) As String
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


'指定フォルダ配下を指定ファイル名で検索してその内容を返す
'読み込んだファイルの内容は1行毎のString配列となるが、
'配列の末尾にはファイルの絶対パスを格納するので注意。
Public Function SearchAndReadFiles(target_folder As String, target_file As String, is_sjis As Boolean) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim file As Object
    For Each file In folder.Files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like target_file Then
            '検索対象のファイルを読み込む
            Dim ts As Object
            If is_sjis = True Then
                Set ts = fso.OpenTextFile(file.path, 1, False, 0)
            Else
                Set ts = fso.OpenTextFile(file.path, 1, False, 1)
            End If
            Dim fileContent As String
            fileContent = ts.ReadAll
            ts.Close
            
            'ファイルの内容を配列に格納して返す
            Dim lines() As String
            lines = Split(fileContent, vbCrLf)
            
            '末尾にファイルパスを追加する
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = file.path
            SearchAndReadFiles = lines
            
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    'サブフォルダも検索する
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        Dim result() As String
        result = SearchAndReadFiles(subFolder.path, target_file, is_sjis)
        If Not IsEmpty(result) Then
            'サブフォルダから結果が返ってきた場合は、その結果を返す
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subFolder
    
    '検索対象のファイルが見つからなかった場合は、空の配列を返す
    SearchAndReadFiles = Split("", vbCrLf)
    Set fso = Nothing
End Function


'現在日時を文字列で返す
Public Function GetNowTimeString() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString = str_date & str_time
End Function

'シートの存在チェック
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

'シートを追加する
Public Sub AddSheet(ByVal sheet_name As String)
    If IsExistSheet(sheet_name) = True Then
        Application.DisplayAlerts = False
        Sheets(sheet_name).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheet_name
End Sub

