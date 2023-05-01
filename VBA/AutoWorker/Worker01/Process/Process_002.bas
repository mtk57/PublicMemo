Attribute VB_Name = "Process_002"
Option Explicit

Private prms As ParamContainer
Private SEP As String
Private DQ As String

Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Dim msg As String: msg = ""

    Set prms = New ParamContainer
    
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    Dim i As Integer
    Dim target As ParamTarget
    Dim targetlist() As ParamTarget
    targetlist = prms.GetTargetList()
        
    WorkerCommon.DoClone prms
    
    For i = LBound(targetlist) To UBound(targetlist)
    
        Set target = targetlist(i)
    
        WorkerCommon.SwitchDevelopBranch prms
        
        WorkerCommon.DoPull prms
        
        CreateFeatureBranch target
        
        DoCopy target
        
        DoAdd
        
        DoCommit target
        
        DoTag target
        
        DoPush target
    
    Next i
        
    Common.WriteLog "Run E"
End Sub

Private Sub CreateFeatureBranch(ByRef target As ParamTarget)
    Common.WriteLog "CreateFeatureBranch S"
    
    Dim cmd As String
    Dim git_result() As String
    
    If WorkerCommon.IsExistBranch(prms, target.GetBranch()) = True Then
        If prms.IsDeleteExistBranch() = False Then
            'featureブランチが既に存在する場合はエラーとする
            Err.Raise 53, , "ブランチが既に存在します。(" & target.GetBranch() & ")"
        Else
            'featureブランチを削除
            cmd = "git branch --delete " & target.GetBranch()
            git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
        End If
    End If
    
    'featureブランチを作成して切り替え
    cmd = "git checkout -b " & target.GetBranch() & " " & prms.GetBaseBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "CreateFeatureBranch E"
End Sub

'VBプロジェクトを収集したフォルダ直下から、コピー元のVBプロジェクトファイルまでのフォルダパスが
'起点フォルダから含めて完全一致するフォルダを見つけて、Gitフォルダにコピーする
Private Sub DoCopy(ByRef target As ParamTarget)
    Common.WriteLog "DoCopy S"
    
    'コピー元のVBプロジェクトファイルまでのフォルダパスを起点フォルダからファイル名までを取得する
    Dim path As String: path = Common.GetStringByKeyword(target.GetVBPrjFilePath(), SEP & prms.GetBaseFolder() & SEP)
    
    '起点フォルダをリネーム
    Dim prj_name As String: prj_name = WorkerCommon.GetProjectName(path)
    path = Replace(path, SEP & prms.GetBaseFolder() & SEP, prms.GetBaseFolder() & "_" & prj_name & SEP)
    
    'VBプロジェクトを収集したフォルダ直下から、一致するファイルがあるかチェックする
    Dim ext As String: ext = Common.GetFileExtension(path)
    Dim file_list() As String: file_list = Common.CreateFileList(prms.GetDestDirPath(), "*." & ext, True)
    
    Dim i As Long
    Dim check_path As String
    Dim is_match As Boolean: is_match = False
    
    For i = LBound(file_list) To UBound(file_list)
        check_path = file_list(i)
        If InStr(check_path, path) > 0 Then
            is_match = True
            Exit For
        End If
    Next i
    
    If is_match = False Then
        Err.Raise 53, , "VBプロジェクトファイルが見つかりません。(" & check_path & ")"
    End If
    
    '起点フォルダをリネームして、Gitフォルダにコピー
    Dim src_path As String: src_path = prms.GetDestDirPath() & SEP & prms.GetBaseFolder() & "_" & prj_name
    
    Dim dst_path As String: dst_path = prms.GetGitDirPath() & SEP
    If ext = "vbp" Then
        dst_path = dst_path & "SRC_020"
    Else
        dst_path = dst_path & "SRC_030"
    End If
    dst_path = dst_path & SEP & prms.GetBaseFolder()
    
    Common.CopyFolder src_path, dst_path
    
    Common.WriteLog "DoCopy E"
End Sub

Private Sub DoAdd()
    Common.WriteLog "DoAdd S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'インデックスに追加する
    cmd = "git add ."
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoAdd E"
End Sub

Private Sub DoCommit(ByRef target As ParamTarget)
    Common.WriteLog "DoCommit S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'コミットする
    cmd = "git commit -m " & DQ & target.GetCommit() & DQ
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoCommit E"
End Sub

Private Sub DoTag(ByRef target As ParamTarget)
    Common.WriteLog "DoTag S"
    
    Dim cmd As String
    Dim git_result() As String
    
    'タグを付ける
    cmd = "git tag " & target.GetTag() & " HEAD"
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoTag E"
End Sub

Private Sub DoPush(ByRef target As ParamTarget)
    Common.WriteLog "DoPush S"
    
    If prms.IsUpdateRemote() = False Then
        Common.WriteLog "DoPush E1"
        Exit Sub
    End If
    
    Dim cmd As String
    Dim git_result() As String
    
    'タグを付ける
    cmd = "git push --tags --set-upstream origin " & target.GetBranch()
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Common.WriteLog "DoPush E"
End Sub



