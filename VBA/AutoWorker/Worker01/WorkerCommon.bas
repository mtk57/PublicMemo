Attribute VB_Name = "WorkerCommon"
Option Explicit

'VBプロジェクトファイルを解析して、参照されているファイルの一覧を取得する
Public Function GetRefFileListForVB6project( _
    ByRef prms As ParamContainer, _
    ByVal vbprj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "GetRefFileListForVB6project S"
    
    Dim ref_files() As String
    
    Dim vbprj_path_new As String
    If InStr(vbprj_path, "/") = 0 Then
        vbprj_path_new = vbprj_path
    Else
        vbprj_path_new = prms.GetGitDirPath() & Application.PathSeparator & Replace(vbprj_path, "/", "\")
    End If

    'vbpファイルに記載されているファイルをリストに追加
    ref_files = ParseVB6Project( _
                    prms, _
                    vbprj_path_new, _
                    contents _
                )
    
    If InStr(vbprj_path, "/") > 0 Then
        '相対パスに変更する
        ref_files = UpdateRefFiles(prms, ref_files)
    End If

    'vbpファイルをリストに追加する
    Dim cnt As Long: cnt = UBound(ref_files)
    ReDim Preserve ref_files(cnt + 1)
    ref_files(cnt + 1) = vbprj_path

    GetRefFileListForVB6project = ref_files
    
    Common.WriteLog "GetRefFileListForVB6project E"
End Function

'VB.NETプロジェクトファイルを解析して、参照されているファイルの一覧を取得する
Public Function GetRefFileListForVBdotNetProject( _
    ByRef prms As ParamContainer, _
    ByVal vbprj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "GetRefFileListForVBdotNetProject S"
    
    Dim ref_files() As String

    Dim vbprj_path_new As String
    If InStr(vbprj_path, "/") = 0 Then
        vbprj_path_new = vbprj_path
    Else
        vbprj_path_new = prms.GetGitDirPath() & Application.PathSeparator & Replace(vbprj_path, "/", "\")
    End If

    'vbprojファイルに記載されているファイルをリストに追加
    ref_files = ParseVBNETProject( _
                    prms, _
                    vbprj_path_new, _
                    contents _
                )
    
    If InStr(vbprj_path, "/") > 0 Then
        '相対パスに変更する
        ref_files = UpdateRefFiles(prms, ref_files)
    End If
    
    'vbprojファイルとslnファイルをリストに追加する
    Dim cnt As Long: cnt = UBound(ref_files)
    ReDim Preserve ref_files(cnt + 2)
    ref_files(cnt + 1) = vbprj_path
    ref_files(cnt + 2) = Replace(vbprj_path, ".vbproj", ".sln")
        
    GetRefFileListForVBdotNetProject = ref_files
    
    Common.WriteLog "GetRefFileListForVBdotNetProject E"
End Function

'相対パスに変更する
Private Function UpdateRefFiles(ByRef prms As ParamContainer, ByRef ref_files() As String) As String()
    Common.WriteLog "UpdateRefFiles S"
    
    Dim ret_files() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    
    For i = LBound(ref_files) To UBound(ref_files)
        ReDim Preserve ret_files(cnt)
        ret_files(cnt) = Replace(Replace(ref_files(i), prms.GetGitDirPath() & Application.PathSeparator, ""), Application.PathSeparator, "/")
        cnt = cnt + 1
    Next i
    
    UpdateRefFiles = ret_files

    Common.WriteLog "UpdateRefFiles E"
End Function

'vbpファイルのパースを行う
' vbp_path : I : vbpファイルパス(絶対パス)
' contents : I : 読み込んだファイルの内容
' Ret : 参照しているファイルのリスト
Public Function ParseVB6Project( _
    ByRef prms As ParamContainer, _
    ByVal vbp_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "ParseVB6Project S"

    Dim i, cnt As Integer
    Dim filelist() As String
    Dim datas() As String
    Dim key As String
    Dim value As String
    
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
        Dim abs_path As String: abs_path = Common.GetAbsolutePathName(base_path, path)
        filelist(cnt) = abs_path
        cnt = cnt + 1
        
        If Common.GetFileExtension(abs_path) = "frm" Then
            Dim frx_path As String: frx_path = Replace(abs_path, ".frm", ".frx")
            If Common.IsExistsFile(frx_path) = True Then
                'frxはvbpに記載されていないのでfrm検知時に存在チェックを行い、存在すればリストに追加する
                 ReDim Preserve filelist(cnt)
                 filelist(cnt) = frx_path
                 cnt = cnt + 1
            End If
        ElseIf Common.GetFileExtension(abs_path) = "ctl" Then
            Dim ctx_path As String: ctx_path = Replace(abs_path, ".ctl", ".ctx")
            If Common.IsExistsFile(ctx_path) = True Then
                'ctxはvbpに記載されていないのでctl検知時に存在チェックを行い、存在すればリストに追加する
                 ReDim Preserve filelist(cnt)
                 filelist(cnt) = ctx_path
                 cnt = cnt + 1
            End If
        End If
        
CONTINUE:
    Next i
    
    ParseVB6Project = filelist
    Common.WriteLog "ParseVB6Project E"
End Function

'vbprojファイルのパースを行う
' vbp_path : I : vbprojファイルパス(絶対パス)
' contents : I : 読み込んだファイルの内容
' Ret : 参照しているファイルのリスト
Public Function ParseVBNETProject( _
    ByRef prms As ParamContainer, _
    ByVal vbproj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "ParseVBNETProject S"

    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    Dim filelist() As String
    
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbproj_path)

    '除外ファイルリストを作成
    Dim ignore_files() As String
    ignore_files = Split(prms.GetIgnoreFiles(), ",")

    cnt = 0

    For i = LBound(contents) To UBound(contents)
        'Common.WriteLog "contents(" & i & ")=" & contents(i)
        
        If InStr(contents(i), "<Compile Include=") = 0 And _
           InStr(contents(i), "<EmbeddedResource Include=") = 0 And _
           InStr(contents(i), "<None Include=") = 0 And _
           InStr(contents(i), "<HintPath>") = 0 And _
           InStr(contents(i), "<ApplicationIcon>") = 0 Then
            'ビルドに必要なファイルを含まないので無視
            'Common.WriteLog "Skip contents(" & i & ")=" & contents(i)
            GoTo CONTINUE
        End If
        
        If Common.IsEmptyArray(ignore_files) = False Then
            For j = LBound(ignore_files) To UBound(ignore_files)
                If InStr(contents(i), ignore_files(j)) > 0 Then
                    '除外ファイルを含むので無視
                    GoTo CONTINUE
                End If
            Next j
        End If
        
        If Common.StartsWith(Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", "")), "packages") Then
            'packagesは無視する
            GoTo CONTINUE
        End If
        
        Dim path As String
        
        If InStr(contents(i), "<Compile Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<Compile Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<EmbeddedResource Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<EmbeddedResource Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<None Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<None Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<HintPath>") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", ""))
        ElseIf InStr(contents(i), "<ApplicationIcon>") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<ApplicationIcon>", ""), "</ApplicationIcon>", ""))
        End If
        
        If path = "" Then
            GoTo CONTINUE
        End If
        
        path = Replace(path, """>", "")
        
        Dim abs_path As String: abs_path = Common.GetAbsolutePathName(base_path, path)
        
        'ビルド対象外ファイルは無視する
        If Common.GetFileExtension(abs_path) <> "vb" And _
           Common.GetFileExtension(abs_path) <> "resx" And _
           Common.GetFileExtension(abs_path) <> "config" And _
           Common.GetFileExtension(abs_path) <> "ico" Then
           
           Common.WriteLog "[ParseVBNETProject] ★★ビルド対象外ファイルが記載されています。確認してください。" & _
                           "【vbproj】=" & vbproj_path & _
                           ", " & _
                           "【ビルド対象外ファイル】=" & abs_path
           GoTo CONTINUE
        End If
        
        '絶対パスに変換する
        ReDim Preserve filelist(cnt)
        filelist(cnt) = abs_path
        cnt = cnt + 1
        
        'ActiveReport 特殊処理
        If InStr(contents(i), "<Compile Include=""reports\") > 0 Then
            'rpxの存在チェックを行い、あれば追加する
            Dim rpx_path As String: rpx_path = Replace(path, ".vb", ".rpx")
            Dim rpx_find_path As String: rpx_find_path = base_path & Application.PathSeparator & rpx_path
            If Common.IsExistsFile(rpx_find_path) = True Then
                Common.WriteLog "rpx found.(" & rpx_find_path & ")"
                
                ReDim Preserve filelist(cnt)
                filelist(cnt) = Common.GetAbsolutePathName(base_path, rpx_path)
                cnt = cnt + 1
            Else
                Common.WriteLog "rpx not found.(" & rpx_find_path & ")"
            End If
        End If
        
CONTINUE:
    Next i
    
    ParseVBNETProject = filelist
    Common.WriteLog "ParseVBNETProject E"
End Function

'ファイルの内容を取得する
Public Function DoShow( _
    ByRef prms As ParamContainer, _
    ByVal path As String _
) As String()
    Common.WriteLog "DoShow S"

    Dim repo_path As String: repo_path = prms.GetGitDirPath()

    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git show " & path
    git_result = Common.RunGit(repo_path, cmd)
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    DoShow = git_result

    Common.WriteLog "DoShow E"
End Function

'タグをファイル名で検索して存在すればファイルパスを返す
Public Function GetFilepathByTag( _
    ByRef prms As ParamContainer, _
    ByVal tag As String, _
    ByVal filename As String _
) As String
    Common.WriteLog "GetFilepathByTag S"

    Dim repo_path As String: repo_path = prms.GetGitDirPath()

    Dim cmd As String
    Dim git_result() As String
    
    'cmd = "git ls-tree -r --name-only " & tag & " | findstr " & filename  'findstrはSJISのみなので不採用
    cmd = "git ls-tree -r --name-only " & tag
    
On Error Resume Next
    git_result = Common.RunGit(repo_path, cmd)
    
    Dim err_msg As String: err_msg = Err.Description
    Err.Clear
On Error GoTo 0

    If err_msg = "" Then
        '成功
    ElseIf InStr(err_msg, "exit code=1") = 0 Then
        'exit code=1以外は上位に再度エラー通知
        Err.Raise 53, , "[GetFilepathByTag] git ls-treeでエラー (err_msg=" & err_msg & ")"
    Else
        'exit code=1はfilenameが見つからない場合と思われるので準正常として動作
        Common.WriteLog "File not found.(tag=" & tag & ", filename=(" & filename & ")"
        GetFilepathByTag = ""
        Exit Function
    End If
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    If Common.IsEmptyArray(git_result) = True Then
        Common.WriteLog "File not found.(tag=" & tag & ", filename=(" & filename & ")"
        GetFilepathByTag = ""
    Else
        '
        Dim find_row_num  As Long: find_row_num = Common.FindRowByKeywordFromArray(filename, git_result, False)
        
        If find_row_num < 0 Then
            '見つからない
            Common.WriteLog "File not found.(tag=" & tag & ", filename=(" & filename & ")"
            GetFilepathByTag = ""
            Exit Function
        End If
    
        If Common.GetFileExtension(filename) = Common.GetFileExtension(git_result(find_row_num)) Then
            GetFilepathByTag = git_result(find_row_num)
        Else
            GetFilepathByTag = ""
        End If
    End If

    Common.WriteLog "GetFilepathByTag E"
End Function

'ローカル/リモートブランチの存在チェック
Public Function IsExistBranch( _
    ByRef prms As ParamContainer, _
    ByVal branch As String _
) As Boolean
    Common.WriteLog "IsExistBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    Dim is_exist_local As Boolean: is_exist_local = IsExistLocalBranch(prms, branch)
    Dim is_exist_remote As Boolean: is_exist_remote = IsExistRemoteBranch(prms, branch)
    
    IsExistBranch = False
    If is_exist_local = True Or is_exist_remote = True Then
         IsExistBranch = True
    End If
    
    Common.WriteLog "IsExistBranch E"
End Function

'ローカルブランチの存在チェック
Public Function IsExistLocalBranch( _
    ByRef prms As ParamContainer, _
    ByVal branch As String _
) As Boolean
    Common.WriteLog "IsExistLocalBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git branch --list " & branch
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistLocalBranch = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistLocalBranch = True
    End If
    
    Common.WriteLog "IsExistLocalBranch E"
End Function

'リモートブランチの存在チェック
Public Function IsExistRemoteBranch( _
    ByRef prms As ParamContainer, _
    ByVal branch As String _
) As Boolean
    Common.WriteLog "IsExistRemoteBranch S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git branch --list --remote origin/" & branch
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistRemoteBranch = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistRemoteBranch = True
    End If
    
    Common.WriteLog "IsExistRemoteBranch E"
End Function

'タグの存在チェック
Public Function IsExistTag( _
    ByRef prms As ParamContainer, _
    ByVal tag As String, _
    ByVal is_local As Boolean _
) As Boolean
    Common.WriteLog "IsExistTag S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    If is_local = True Then
        cmd = "git tag --list " & tag
    Else
        cmd = "git ls-remote --tags origin " & tag
    End If
    
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistTag = False
    If Common.IsEmptyArray(git_result) = False Then
        
        Dim i As Long
        For i = 0 To UBound(git_result)
            If InStr(git_result(i), tag) > 0 Then
                IsExistTag = True
                Exit For
            End If
        Next i
        
    End If
    
    Common.WriteLog "IsExistTag E"
End Function

'VBプロジェクト名を返す
Public Function GetProjectName(ByVal vbprj_file_path As String) As String
    Common.WriteLog "GetProjectName S"
    
    Dim vbprj_file_name As String: vbprj_file_name = Common.GetFileName(vbprj_file_path)
    Dim ext As String: ext = Common.GetFileExtension(vbprj_file_name)
    
    GetProjectName = Replace(vbprj_file_name, "." & ext, "")
    
    Common.WriteLog "GetProjectName E"
End Function

Public Sub DoClone(ByRef prms As ParamContainer)
    Common.WriteLog "DoClone S"
    
    If prms.GetUrl() = "" Then
        Common.WriteLog "DoClone E1"
        Exit Sub
    End If
    
    If Common.IsExistsFolder(prms.GetGitDirPath()) = True Then
        Common.WriteLog "DoClone E2"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    'クローン
    cmd = "git clone " & prms.GetUrl() & " " & prms.GetGitDirPath()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoClone E"
End Sub

Public Sub SwitchDevelopBranch(ByRef prms As ParamContainer)
    Common.WriteLog "SwitchDevelopBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'developブランチが存在しない場合はエラーとする
    If IsExistBranch(prms, prms.GetBaseBranch()) = False Then
        Err.Raise 53, , "[SwitchDevelopBranch] ブランチが存在しません。(" & prms.GetBaseBranch() & ")"
    End If
    
    'カレントブランチを確認する
    cmd = "git branch --show-current"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If git_result(0) = prms.GetBaseBranch() Then
        Common.WriteLog "SwitchDevelopBranch E1"
        Exit Sub
    End If
    
    'ローカルブランチがあるか確認する
    Dim is_exist_local As Boolean: is_exist_local = False
    
    cmd = "git branch --list " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If Common.IsEmptyArray(git_result) = False Then
        is_exist_local = True
    End If
    
    If is_exist_local = True Then
        'ローカルブランチがあるのでswitchで切り替え
        cmd = "git switch " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    Else
        'ローカルブランチがないのでリモートから作成して切り替え
        cmd = "git checkout -b " & prms.GetBaseBranch() & " origin/" & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If
    
    Common.WriteLog "SwitchDevelopBranch E"
End Sub

Public Sub DoPull(ByRef prms As ParamContainer)
    Common.WriteLog "DoPull S"
    
    If prms.IsUpdateRemote() = False Then
        Common.WriteLog "DoPull E1"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    'プル
    cmd = "git pull origin " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPull E"
End Sub

Public Function DoPush(ByRef prms As ParamContainer, ByVal cmd As String) As String()
    Common.WriteLog "DoPush S"
    
    Dim git_result() As String
    
    If prms.IsUpdateRemote() = False Then
        DoPush = git_result
        Common.WriteLog "DoPush E1"
        Exit Function
    End If
    
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    DoPush = git_result
    
    Common.WriteLog "DoPush E"
End Function

Public Sub SwitchBranch(ByRef prms As ParamContainer, ByVal target As ParamTarget)
    Common.WriteLog "SwitchBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    '対象ブランチが存在しない場合はエラーとする
    If IsExistBranch(prms, target.GetBranch()) = False Then
        Err.Raise 53, , "[SwitchBranch] ブランチが存在しません。(" & target.GetBranch() & ")"
    End If
    
    'カレントブランチを確認する
    cmd = "git branch --show-current"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If git_result(0) = target.GetBranch() Then
        Common.WriteLog "SwitchBranch E1"
        Exit Sub
    End If
    
    'ローカルブランチがあるか確認する
    Dim is_exist_local As Boolean: is_exist_local = False
    
    cmd = "git branch --list " & target.GetBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If Common.IsEmptyArray(git_result) = False Then
        is_exist_local = True
    End If
    
    If is_exist_local = True Then
        'ローカルブランチがあるのでswitchで切り替え
        cmd = "git switch " & target.GetBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    Else
        'ローカルブランチがないのでリモートから作成して切り替え
        cmd = "git checkout -b " & target.GetBranch() & " origin/" & target.GetBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If
    
    Common.WriteLog "SwitchBranch E"
End Sub

'targetに以下のいずれかを指定する
'1.マージ元のブランチ名
'2.マージ元のタグ名
Public Sub DoMerge(ByRef prms As ParamContainer, ByVal target As String)
    Common.WriteLog "DoMerge S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'マージ
    cmd = "git merge " & target
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoMerge E"
End Sub


