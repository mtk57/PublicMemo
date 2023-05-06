Attribute VB_Name = "WorkerCommon"
Option Explicit

'vbpファイルのパースを行う
' vbp_path : I : vbpファイルパス(絶対パス)
' contents : I : 読み込んだファイルの内容
' Ret : 参照しているファイルのリスト
Public Function ParseVB6Project( _
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
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
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
    ByVal vbproj_path As String, _
    ByRef contents() As String _
) As String()
    Common.WriteLog "ParseVBNETProject S"

    Dim i, cnt As Integer
    Dim filelist() As String
    
    Dim base_path As String: base_path = Common.GetFolderNameFromPath(vbproj_path)

    cnt = 0

    For i = LBound(contents) To UBound(contents)
        If InStr(contents(i), "<Compile Include=") = 0 And _
           InStr(contents(i), "<EmbeddedResource Include=") = 0 And _
           InStr(contents(i), "<None Include=") = 0 And _
           InStr(contents(i), "<HintPath>") = 0 Then
            'ビルドに必要なファイルを含まないので無視
            GoTo CONTINUE
        End If
        
        ReDim Preserve filelist(cnt)
        
        Dim path As String
        
        If InStr(contents(i), "<Compile Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<Compile Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<EmbeddedResource Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<EmbeddedResource Include=""", ""), """ />", ""))
        ElseIf InStr(contents(i), "<None Include=") > 0 Then
            path = Trim(Replace(Replace(contents(i), "<None Include=""", ""), """ />", ""))
        Else
            path = Trim(Replace(Replace(contents(i), "<HintPath>", ""), "</HintPath>", ""))
        End If
        
        path = Replace(path, """>", "")
        
        '絶対パスに変換する
        filelist(cnt) = Common.GetAbsolutePathName(base_path, path)
        cnt = cnt + 1
        
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

'指定ファイル名がタグに含まれるか検索して、存在すればファイルパスを返す
Public Function GetFilePathByTag( _
    ByRef prms As ParamContainer, _
    ByVal tag As String, _
    ByVal filename As String _
) As String
    Common.WriteLog "GetFilePathByTag S"
    
    Dim git_result() As String
    
    'タグに含まれれるファイルの一覧を取得する
    git_result = DoLsTree(prms, tag)
    
    If Common.IsEmptyArray(git_result) = True Then
        Common.WriteLog "File not found. (tag=" & tag & ")"
        Common.WriteLog "GetFilePathByTag E1"
        GetFilePathByTag = ""
        Exit Function
    End If
    
    Dim i As Long
    
    For i = LBound(git_result) To UBound(git_result)
        If InStr(git_result(i), "/" & filename) > 0 Then
            If Common.GetFileExtension(git_result(i)) = Common.GetFileExtension(filename) Then
                GetFilePathByTag = git_result(i)
                Common.WriteLog "GetFilePathByTag E2"
                Exit Function
            End If
        End If
    Next i
    
    Common.WriteLog "File not found. (filename=" & filename & ")"
    GetFilePathByTag = ""

    Common.WriteLog "GetFilePathByTag E"
End Function

'タグに含まれれるファイルの一覧を取得する
Public Function DoLsTree( _
    ByRef prms As ParamContainer, _
    ByVal tag As String _
) As String()
    Common.WriteLog "DoLsTree S"

    Dim repo_path As String: repo_path = prms.GetGitDirPath()

    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git ls-tree -r --name-only " & tag
    git_result = Common.RunGit(repo_path, cmd)
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    DoLsTree = git_result

    Common.WriteLog "DoLsTree E"
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
    ByVal tag As String _
) As Boolean
    Common.WriteLog "IsExistTag S"
    
    Dim repo_path As String: repo_path = prms.GetGitDirPath()
    
    Dim cmd As String
    Dim git_result() As String
    
    cmd = "git tag --list " & tag
    git_result = Common.RunGit(repo_path, cmd)
    git_result = Common.DeleteEmptyArray(git_result)
    
    IsExistTag = False
    If Common.IsEmptyArray(git_result) = False Then
        IsExistTag = True
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
        Err.Raise 53, , "ブランチが存在しません。(" & prms.GetBaseBranch() & ")"
    End If
    
    'カレントブランチを確認する
    cmd = "git branch --show-current"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    If git_result(0) = prms.GetBaseBranch() Then
        Common.WriteLog "CheckoutBranch E1"
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
        'ローカルブランチがないので作成して切り替え
        cmd = "git checkout -b " & prms.GetBaseBranch()
        git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    End If
    
    Common.WriteLog "SwitchDevelopBranch E"
End Sub

Public Sub DoPull(ByRef prms As ParamContainer)
    Common.WriteLog "DoPull S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'プル
    cmd = "git pull origin " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPull E"
End Sub


