Attribute VB_Name = "Process"
Option Explicit

'定数
Private Const MAIN_SHEET = "main"
Private Const SEARCH_FILE_NAME = "O5"
Private Const SEARCH_DIR_PATH = "O6"
Private Const FILE_ENCODE = "O7"
Private Const OUT_DIR_PATH = "O11"
Private Const OUT_SHEET_NAME = "O12"
Private Const OUT_BAT_PATH = "O13"

'シートから収集した情報
Private search_file As String
Private search_path As String
Private encode As String
Private out_path As String
Private out_sheet As String
Private out_bat As String

'メイン処理
Public Sub Run()
    Worksheets(MAIN_SHEET).Activate
    Dim err_msg As String

    'mainシートの情報を収集
    search_file = Range(SEARCH_FILE_NAME).value
    search_path = Range(SEARCH_DIR_PATH).value
    encode = Range(FILE_ENCODE).value
    out_path = Range(OUT_DIR_PATH).value
    out_sheet = Range(OUT_SHEET_NAME).value
    out_bat = Range(OUT_BAT_PATH).value
    
    '収集した情報を検証する
    err_msg = Validate()
    If err_msg <> "" Then
        MsgBox err_msg
        Exit Sub
    End If
    
    'ファイルエンコード
    Dim is_sjis As Boolean: is_sjis = True
    If encode = "UTF-8" Then
        is_sjis = False
    End If
    
    'VBプロジェクトファイルを検索して読み込む
    Dim contents() As String: contents = Common.SearchAndReadFiles(search_path, search_file, is_sjis)
    
    If UBound(contents) = -1 Then
        MsgBox "VBプロジェクトファイルが見つかりませんでした"
        Exit Sub
    End If
    
    'VBプロジェクトファイルのパースを行う
    Dim filelist() As String: filelist = ParseContents(contents, search_file)
    
    'VBプロジェクトファイルが参照しているファイルを同じフォルダ構成のままコピーする
    CopyProjectFiles out_path, filelist
    
    'BATファイルを作成する
    CreateBatFile search_path, out_path, out_bat, filelist
    
    'シート名が指定されていればシートにVBプロジェクトファイルを出力する
    'TODO

End Sub

'収集した情報を検証する
Private Function Validate() As String
    If search_file = "" Or _
       search_path = "" Or _
       encode = "" Or _
       out_path = "" Then
        Validate = "未入力の情報があります"
        Exit Function
    End If

    Dim ext As String: ext = Common.GetFileExtension(search_file)
    
    If ext <> "vbp" And ext <> "vbproj" Then
        Validate = "VBプロジェクトファイル名が未対応の拡張子です"
        Exit Function
    End If

    If Common.IsExistsFolder(search_path) = False Then
        Validate = "検索フォルダが存在しません"
        Exit Function
    End If
    
    If out_bat <> "" Then
        ext = Common.GetFileExtension(out_bat)
        If ext <> "bat" Then
            Validate = "BATファイル名が未対応の拡張子です"
            Exit Function
        End If
    End If

    Validate = ""
End Function

'VBプロジェクトファイルのパースを行う
Private Function ParseContents(ByRef contents() As String, ByVal filename As String) As String()
    Dim ext As String: ext = Common.GetFileExtension(filename)
    
    If ext = "vbp" Then
        ParseContents = ParseVB6Project(contents)
    Else
        ParseContents = ParseVBNETProject(contents)
    End If

End Function

'vbpファイルのパースを行う
'
'vbpファイルのパース対象と内容の例は以下の通り。
'-----------------------------------------
'Module=module1; module1.bas
'Module=module2; ..\cmn\module2.bas
'Module=module3; sub\module3.bas
'Form=form1.frm
'Form=..\cmn\form2.frm
'Form=sub\form3.frm
'Class=class1; class1.cls
'Class=class2; ..\cmn\class2.cls
'Class=class3; sub\class3.cls
'ResFile32="resfile321.RES"
'ResFile32="..\cmn\resfile322.RES"
'ResFile32="sub\resfile323.RES"
'-----------------------------------------
'上記例の場合、以下の配列が返る (base_pathがC:\tmp\baseの場合)
'[0] : "C:\tmp\base\module1.bas"
'[1] : "C:\tmp\cmn\module2.bas"
'[2] : "C:\tmp\base\sub\module3.bas"
'[3] : "C:\tmp\base\form1.frm"
'[4] : "C:\tmp\cmn\form2.frm"
'[5] : "C:\tmp\base\sub\form3.frm"
'[6] : "C:\tmp\base\class1.cls"
'[7] : "C:\tmp\cmn\class2.cls"
'[8] : "C:\tmp\base\sub\class3.cls"
'[9] : "C:\tmp\base\resfile321.RES"
'[10] :"C:\tmp\cmn\resfile322.RES"
'[11] :"C:\tmp\base\sub\resfile323.RES"
'[12] :"C:\tmp\base\test.vbp"
Private Function ParseVB6Project(ByRef contents() As String) As String()
    Dim i As Integer
    Dim filelist() As String
    Dim datas() As String
    Dim key As String
    Dim value As String
    
    Dim vbp_path As String: vbp_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbp_path)

    For i = LBound(contents) To UBound(contents)
        If InStr(contents(i), "=") = 0 Then
            '"="を含まないので無視
            GoTo CONTINUE
        End If
        
        'Key/Valueに分ける
        datas = Split(contents(i), "=")
        
        'キーを取得
        key = datas(0)
        
        '対象キーか?
        If key <> "Module" And key <> "Form" And key <> "Class" And key <> "ResFile32" Then
            '対象外なので無視
            GoTo CONTINUE
        End If
        
        '値を取得
        value = Replace(datas(1), """", "")
        
        ReDim Preserve filelist(i)
        Dim path As String
        
        If InStr(value, ";") > 0 Then
            path = Trim(Split(value, ";")(1))
        Else
            path = Trim(value)
        End If
        
        '絶対パスに変換する
        filelist(i) = Common.GetAbsolutePathName(base_path, path)
        
CONTINUE:
    Next i
    
    '最後にvbpファイルコピーも追加する
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 1)
    filelist(filelist_cnt + 1) = vbp_path
    
    ParseVB6Project = filelist
End Function

'vbprojファイルのパースを行う
'
'vbprojファイルのパース対象と内容の例は以下の通り。
'-----------------------------------------
'TBD
'-----------------------------------------
Private Function ParseVBNETProject(ByRef contents() As String) As String()
    'TBD
    ParseVBNETProject = Nothing
End Function

'VBプロジェクトファイルが参照しているファイルを同じフォルダ構成のままコピーする
Private Sub CopyProjectFiles(ByVal dest_path As String, ByRef filelist() As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim src_base_path As String: src_base_path = Common.GetCommonString(filelist)
    Dim i As Integer
    
    For i = LBound(filelist) To UBound(filelist)
        Dim src_path As String: src_path = filelist(i)
        Dim dst_path As String: dst_path = Replace(src_path, src_base_path, dest_path & Application.PathSeparator)
        Dim path As String: path = Common.GetFolderNameFromPath(dst_path)
        
        'フォルダが存在しない場合は作成する
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        'ファイルをコピーする
        fso.CopyFile src_path, dst_path
    Next i
    
    Set fso = Nothing
End Sub

'BATファイルを作成する
'作成イメージ (SJISで作成すること)
'-------------------
'@echo off
'set SRC_DIR=C:\src
'set DST_DIR=C:\dst
'
'echo SRC_DIR=%SRC_DIR%
'echo DST_DIR=%DST_DIR%
'
'REM フォルダ階層だけコピー
'xcopy /E /I /Y /T "%SRC_DIR%" "%DST_DIR%"
'
'REM 各ファイルをコピー
'xcopy /Y /F "%SRC_DIR%\base\module1.bas"        "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\module2.bas"         "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\module3.bas"    "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\form1.frm"          "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\form2.frm"           "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\form3.frm"      "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\class1.cls"         "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\class2.cls"          "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\class3.cls"     "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\resfile321.RES"     "%DST_DIR%\base"
'xcopy /Y /F "%SRC_DIR%\cmn\resfile322.RES"      "%DST_DIR%\cmn"
'xcopy /Y /F "%SRC_DIR%\base\sub\resfile323.RES" "%DST_DIR%\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\test.vbp"           "%DST_DIR%\base"
'
'pause
'-------------------
Private Sub CreateBatFile(ByVal src_path As String, ByVal dst_path As String, ByVal bat_path As String, ByRef filelist() As String)
    Dim i As Integer
    Dim contents() As String
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    Dim file, src, dst As String
    
    Const OFFSET = 13
    ReDim Preserve contents(filelist_cnt + OFFSET)
    
    'コマンド作成開始
    contents(0) = "@echo off"
    contents(1) = "set SRC_DIR=" & src_path
    contents(2) = "set DST_DIR=" & dst_path
    contents(3) = ""
    contents(4) = "echo SRC_DIR=%SRC_DIR%"
    contents(5) = "echo DST_DIR=%DST_DIR%"
    contents(6) = ""
    contents(7) = "REM フォルダ階層だけコピー"
    contents(8) = "xcopy /E /I /Y /T ""%SRC_DIR%"" ""%DST_DIR%"""
    contents(9) = ""
    contents(10) = "REM 各ファイルをコピー"

    For i = LBound(filelist) To UBound(filelist)
        file = filelist(i)
        src = "%SRC_DIR%" & Replace(file, src_path, "")
        dst = "%DST_DIR%" & Replace(Common.GetFolderNameFromPath(file), src_path, "")
        contents(i + OFFSET - 2) = "xcopy /Y /F " & """" & src & """" & " " & """" & dst & """"
    Next i
    
    contents(filelist_cnt + OFFSET - 1) = ""
    contents(filelist_cnt + OFFSET) = "pause"
    
    'ファイルに出力する
    Common.CreateSJISTextFile contents, bat_path

End Sub


