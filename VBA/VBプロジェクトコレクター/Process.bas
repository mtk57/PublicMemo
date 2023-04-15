Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Private main_param As MainParam
Private sub_param As SubParam

Private vbprj_files() As String


'メイン処理
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    'パラメータのチェックと収集を行う
    If CheckAndCollectParam() = False Then
        Common.WriteLog "Run E1"
        Exit Sub
    End If
    
    'VBプロジェクトファイルを検索する
    If SearchVBProjFile() = False Then
        Common.WriteLog "Run E2"
        Exit Sub
    End If
    
    Dim i As Integer
    Dim copy_files() As String
    
    'メインループ
    For i = LBound(vbprj_files) To UBound(vbprj_files)
        Dim vbproj_path As String: vbproj_path = vbprj_files(i)
        Common.WriteLog "i=" & i & ":[" & vbproj_path & "]"
    
        'VBプロジェクトファイルのパースを行い、コピーするファイルリストを作成する
        copy_files = CreateCopyFileList(vbproj_path)
        
        'VBプロジェクトファイルが参照しているファイルを同じフォルダ構成のままコピーする
        Dim dst_path As String: dst_path = main_param.GetDestDirPath() & SEP & GetProjectName(vbproj_path)
        CopyProjectFiles dst_path, copy_files
        
        'BATファイルを作成する
        CreateBatFile vbproj_path, dst_path, copy_files
    
        'VBプロジェクトファイルをシート出力する
        OutputSheet vbproj_path
    Next i

    Common.WriteLog "Run E"
    MsgBox "終わりました"
End Sub

'パラメータのチェックと収集を行う
Private Function CheckAndCollectParam() As Boolean
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    err_msg = main_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        Common.WriteLog "CheckAndCollectParam E1 (" & err_msg & ")"
        CheckAndCollectParam = False
        Exit Function
    End If
    
    Common.WriteLog main_param.GetAllValue()

    'Sub Params
    Set sub_param = New SubParam
    err_msg = sub_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        Common.WriteLog "CheckAndCollectParam E2 (" & err_msg & ")"
        CheckAndCollectParam = False
        Exit Function
    End If
    
    'Main Param、Sub ParamのどちらにもVBプロジェクトファイルが指定されていない場合はNG
    If main_param.GetVBPrjFileName() = "" And _
       sub_param.GetVBProjFilePathListCount() <= 0 Then
       err_msg = "VBプロジェクトファイルが指定されていません。"
        MsgBox err_msg
        Common.WriteLog "CheckAndCollectParam E3 (" & err_msg & ")"
        CheckAndCollectParam = False
        Exit Function
    End If

    CheckAndCollectParam = True
    Common.WriteLog "CheckAndCollectParam E"
End Function

'VBプロジェクトファイルを検索する
Private Function SearchVBProjFile() As Boolean
    Common.WriteLog "SearchVBProjFile S"
    
    Dim err_msg As String
    Dim path As String
    Dim i As Integer: i = 0
    
    'VBプロジェクトファイルを検索する
    If main_param.GetVBPrjFileName() <> "" Then
        path = Common.SearchFile(main_param.GetSrcDirPath(), main_param.GetVBPrjFileName())
        ReDim Preserve vbprj_files(i)
        vbprj_files(i) = path
    End If
    
    'Sub Paramに指定されたパスをマージ
    If sub_param.GetVBProjFilePathListCount() > 0 Then
        vbprj_files = Common.MergeArray(vbprj_files, sub_param.GetVBProjFilePathList())
    End If
    
    vbprj_files = Common.DeleteEmptyArray(vbprj_files)
    
    If Common.IsEmptyArray(vbprj_files) = True Then
        err_msg = "VBプロジェクトファイルが見つかりませんでした"
        MsgBox err_msg
        Common.WriteLog "SearchVBProjFile E1 (" & err_msg & ")"
        SearchVBProjFile = False
        Exit Function
    End If
    
    SearchVBProjFile = True
    Common.WriteLog "SearchVBProjFile E"
End Function

'VBプロジェクトファイルのパースを行い、コピーするファイルリストを取得する
Private Function CreateCopyFileList(ByVal vbproj_path As String) As String()
    Common.WriteLog "CreateCopyFileList S"
    
    'VBプロジェクトファイルのパースを行う
    CreateCopyFileList = ParseContents(vbproj_path)
    
    Common.WriteLog "CreateCopyFileList E"
End Function

'VBプロジェクトファイルのパースを行う
Private Function ParseContents(ByVal vbproj_path As String) As String()
    Common.WriteLog "ParseContents S"
    
    'VBプロジェクトファイルの内容を読み込む
    Dim contents() As String: contents = GetVBPrjContents(vbproj_path)
    
    '末尾にファイルパスを追加する
    Dim cnt As Integer: cnt = UBound(contents)
    ReDim Preserve contents(cnt + 1)
    contents(cnt + 1) = vbproj_path
    
    If Common.GetFileExtension(vbproj_path) = "vbp" Then
        ParseContents = ParseVB6Project(contents)
    Else
        ParseContents = ParseVBNETProject(contents)
    End If

    Common.WriteLog "ParseContents E"
End Function

'VBプロジェクトファイルの内容を読み込む
Private Function GetVBPrjContents(ByVal vbproj_path As String) As String()
    Common.WriteLog "GetVBPrjContents S"
    
    'VBプロジェクトファイルの内容を読み込む
    Dim raw_contents As String: raw_contents = Common.ReadTextFileBySJIS(vbproj_path)
    
    'ファイルの内容を配列に格納する
    Dim contents() As String: contents = Split(raw_contents, vbCrLf)
    
    GetVBPrjContents = Common.DeleteEmptyArray(contents)
    
    Common.WriteLog "GetVBPrjContents E"
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
'UserControl = usercontrol1.ctl
'UserControl=..\cmn\usercontrol2.ctl
'UserControl=sub\usercontrol3.ctl
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
'[12] :"C:\tmp\base\usercontrol1.ctl
'[13] :"C:\tmp\cmn\usercontrol2.ctl
'[14] :"C:\tmp\base\sub\usercontrol3.ctl
'[15] :"C:\tmp\base\test.vbp"
Private Function ParseVB6Project(ByRef contents() As String) As String()
    Common.WriteLog "ParseVB6Project S"

    Dim i, cnt As Integer
    Dim filelist() As String
    Dim datas() As String
    Dim key As String
    Dim value As String
    
    Dim vbp_path As String: vbp_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbp_path)

    cnt = 0

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
        If key <> "Module" And key <> "Form" And key <> "Class" And key <> "ResFile32" And key <> "UserControl" Then
            '対象外なので無視
            GoTo CONTINUE
        End If
        
        '値を取得
        value = Replace(datas(1), """", "")
        
        ReDim Preserve filelist(cnt)
        Dim path As String
        
        If InStr(value, ";") > 0 Then
            path = Trim(Split(value, ";")(1))
        Else
            path = Trim(value)
        End If
        
        '絶対パスに変換する
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
CONTINUE:
    Next i
    
    '最後にvbpファイルも追加する
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 1)
    filelist(filelist_cnt + 1) = vbp_path
    
    ParseVB6Project = filelist
    Common.WriteLog "ParseVB6Project E"
End Function

'vbprojファイルのパースを行う
'
'vbprojファイルのパース対象と内容の例は以下の通り。
'-----------------------------------------
'<?xml version="1.0" encoding="utf-8"?>
'<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
'  <ItemGroup>
'    <Compile Include="..\cmn\cmn.vb" />
'    <Compile Include="base.vb" />
'    <Compile Include="sub\sub.vb" />
'  </ItemGroup>
'</Project>
'-----------------------------------------
'上記例の場合、以下の配列が返る (base_pathがC:\tmp\baseの場合)
'[0] : "C:\tmp\base\base.vb"
'[1] : "C:\tmp\cmn\cmn.vb"
'[2] : "C:\tmp\base\sub\sub.vb"
'[3] : "C:\tmp\base\test.vbproj"
Private Function ParseVBNETProject(ByRef contents() As String) As String()
    Common.WriteLog "ParseVBNETProject S"

    Dim i, cnt As Integer
    Dim filelist() As String
    
    Dim vbproj_path As String: vbproj_path = contents(UBound(contents))
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbproj_path)

    cnt = 0

    For i = LBound(contents) To UBound(contents)
        'Dim find_word As String: find_word = ".vb" & """ />"
        'If InStr(contents(i), find_word) = 0 Then
        '    '".vb" />"を含まないので無視
        '    GoTo CONTINUE
        'End If
        
        If InStr(contents(i), "<Compile Include=") = 0 And _
           InStr(contents(i), "<EmbeddedResource Include=") = 0 And _
           InStr(contents(i), "<None Include=") = 0 Then
            'ビルドに必要なファイルを含まないので無視
            GoTo CONTINUE
        End If
        
        ReDim Preserve filelist(cnt)
        
        Dim path As String
        
        If InStr(contents(i), "<Compile Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<Compile Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<EmbeddedResource Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<EmbeddedResource Include=""", ""), """ />", ""))
        Else
            path = Trim(Replace(Replace(contents(i), "<None Include=""", ""), """ />", ""))
        End If
        
        path = Replace(path, """>", "")
        
        '絶対パスに変換する
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
CONTINUE:
    Next i
    
    '最後にvbproj, slnファイルも追加する
    Dim filelist_cnt As Integer: filelist_cnt = UBound(filelist)
    ReDim Preserve filelist(filelist_cnt + 2)
    filelist(filelist_cnt + 1) = vbproj_path
    filelist(filelist_cnt + 2) = Replace(vbproj_path, ".vbproj", ".sln")
    
    ParseVBNETProject = filelist
    Common.WriteLog "ParseVBNETProject E"
End Function

'VBプロジェクトファイルが参照しているファイルを同じフォルダ構成のままコピーする
Private Sub CopyProjectFiles(ByVal in_dest_path As String, ByRef filelist() As String)
    Common.WriteLog "CopyProjectFiles S"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim SEP As String: SEP = Application.PathSeparator
    Dim base_path As String: base_path = Common.GetCommonString(filelist)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim i As Integer
    
    For i = LBound(filelist) To UBound(filelist)
        Dim src As String: src = filelist(i)
        Dim dst As String: dst = in_dest_path & SEP & dst_base_path & Replace(src, base_path, "")
        Dim path As String: path = Common.GetFolderNameFromPath(dst)
        
        'フォルダが存在しない場合は作成する
        If Not fso.FolderExists(path) Then
            Common.CreateFolder (path)
        End If
        
        If Right(src, 4) = ".sln" And _
           Common.IsExistsFile(src) = False Then
           'slnの場合、コピー元に存在しない場合
           GoTo CONTINUE
        End If
        
        'ファイルをコピーする
        fso.CopyFile src, dst
        
CONTINUE:
        
    Next i
    
    Set fso = Nothing
    Common.WriteLog "CopyProjectFiles E"
End Sub

Private Function GetProjectName(ByVal vbprj_file_path As String) As String
    Common.WriteLog "GetProjectName S"
    Dim vbprj_file_name As String: vbprj_file_name = Common.GetFileName(vbprj_file_path)
    Dim ext As String: ext = Common.GetFileExtension(vbprj_file_name)
    GetProjectName = Replace(vbprj_file_name, "." & ext, "")
    Common.WriteLog "GetProjectName E"
End Function

'BATファイルを作成する
'作成イメージ (SJISで作成すること)
'-------------------
'@echo off
'set SRC_DIR=C:\src
'set DST_DIR=C:\_tmp
'
'echo SRC_DIR=%SRC_DIR%
'echo DST_DIR=%DST_DIR%
'
'REM 各ファイルをコピー
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\module1.bas" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\module2.bas" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\module3.bas" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\form1.frm" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\form2.frm" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\form3.frm" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\class1.cls" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\class2.cls" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\class3.cls" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\resfile321.RES" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\resfile322.RES" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\resfile323.RES" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\usercontrol1.ctl" "%DST_DIR%\C\src\base"
'
'md "%DST_DIR%\C\src\cmn"
'xcopy /Y /F "%SRC_DIR%\cmn\usercontrol2.ctl" "%DST_DIR%\C\src\cmn"
'
'md "%DST_DIR%\C\src\base\sub"
'xcopy /Y /F "%SRC_DIR%\base\sub\usercontrol3.ctl" "%DST_DIR%\C\src\base\sub"
'
'md "%DST_DIR%\C\src\base"
'xcopy /Y /F "%SRC_DIR%\base\test.vbp" "%DST_DIR%\C\src\base"
'
'
'pause
'-------------------
Private Sub CreateBatFile(ByVal vbproj_path As String, ByVal dst_path As String, ByRef copy_files() As String)
    Common.WriteLog "CreateBatFile S"

    If main_param.IsCreateBat() = False Then
        Common.WriteLog "CreateBatFile E1"
        Exit Sub
    End If
    
    Dim i As Integer
    Dim contents() As String
    Dim contents_cnt As Integer
    Dim base_path As String: base_path = Common.GetCommonString(copy_files)
    Dim dst_base_path As String: dst_base_path = Replace(base_path, ":", "")
    Dim bat_name As String: bat_name = GetProjectName(vbproj_path) & ".bat"

    Const FIRST_ROW_CNT = 7
    Const ROW_CNT = 3
    Const SECOND_ROW_CNT = 2
    
    ReDim Preserve contents(FIRST_ROW_CNT)
    
    'コマンド作成開始
    contents(0) = "@echo off"
    contents(1) = "set SRC_DIR=" & Common.RemoveTrailingBackslash(base_path)
    contents(2) = "set DST_DIR=" & dst_path
    contents(3) = ""
    contents(4) = "echo SRC_DIR=%SRC_DIR%"
    contents(5) = "echo DST_DIR=%DST_DIR%"
    contents(6) = ""
    contents(7) = "REM 各ファイルをコピー"
    
    Dim OFFSET As Integer: OFFSET = UBound(contents) + 1

    For i = LBound(copy_files) To UBound(copy_files)
        contents_cnt = UBound(contents)
        ReDim Preserve contents(contents_cnt + ROW_CNT)
    
        Dim file As String: file = copy_files(i)
        
        Dim src As String: src = "%SRC_DIR%" & SEP & Replace(file, base_path, "")
        Dim dst_tmp As String: dst_tmp = "%DST_DIR%" & SEP & dst_base_path & Replace(file, base_path, "")
        Dim dst As String: dst = Common.GetFolderNameFromPath(dst_tmp)
        
        contents(i * ROW_CNT + OFFSET) = "md " & DQ & dst & DQ
        contents(i * ROW_CNT + OFFSET + 1) = "xcopy /Y /F " & DQ & src & DQ & " " & DQ & dst & DQ
        contents(i * ROW_CNT + OFFSET + 2) = ""
    Next i
    
    contents_cnt = UBound(contents)
    ReDim Preserve contents(contents_cnt + SECOND_ROW_CNT)
    contents(contents_cnt + SECOND_ROW_CNT - 1) = ""
    contents(contents_cnt + SECOND_ROW_CNT) = "pause"
    
    'ファイルに出力する
    Common.CreateSJISTextFile contents, dst_path & SEP & bat_name
    
    Common.WriteLog "CreateBatFile E"
End Sub

'VBプロジェクトファイルをシート出力する
Private Sub OutputSheet(ByVal vbproj_path As String)
    Common.WriteLog "OutputSheet S"

    If main_param.IsOutSheet() = False Then
        Common.WriteLog "OutputSheet E1"
        Exit Sub
    End If
    
    Dim sheet_name As String: sheet_name = GetProjectName(vbproj_path)
    
    'VBプロジェクトファイルの内容を読み込む
    Dim contents() As String: contents = GetVBPrjContents(vbproj_path)
    
    Dim prj_path As String: prj_path = contents(UBound(contents))
    
    Dim before_sheet_name As String: before_sheet_name = ActiveSheet.Name
    
    Common.AddSheet sheet_name
    
    'ファイルの内容を指定されたシートに出力する
    Common.OutputTextFileToSheet vbproj_path, sheet_name
    
    ThisWorkbook.Sheets(before_sheet_name).Select
    
    Common.WriteLog "OutputSheet E"
End Sub
