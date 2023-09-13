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
    
    prms.SetProcessType PROCESS_TYPE.PROC_002
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
            Err.Raise 53, , "[CreateFeatureBranch] ブランチが既に存在します。(" & target.GetBranch() & ")"
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
    
    If Common.IsEmptyArray(file_list) = False Then
        For i = LBound(file_list) To UBound(file_list)
            check_path = file_list(i)
            If InStr(check_path, path) > 0 Then
                is_match = True
                Exit For
            End If
        Next i
    End If
    
    If is_match = False Then
        Err.Raise 53, , "[DoCopy] VBプロジェクトファイルが見つかりません。path=(" & prms.GetDestDirPath() & ")"
    End If
    
    '起点フォルダをリネームして、Gitフォルダにコピー
    Dim src_path As String: src_path = prms.GetDestDirPath() & SEP & prms.GetBaseFolder() & "_" & prj_name
    
    Dim dst_path As String: dst_path = prms.GetGitDirPath() & SEP & target.GetDestBaseDir() & SEP & prms.GetBaseFolder()
    
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
    
    'Run_002の場合はSTEP1.1以外はエラーとする
    If InStr(target.GetTag(), "STEP1.1") = 0 Then
        Err.Raise 53, , "[DoTag] STEP1.1が指定されていません。 (tag=" & target.GetTag() & ")"
    End If
    
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
    
On Error Resume Next
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    Dim err_msg As String: err_msg = Err.Description
    Err.Clear
On Error GoTo 0

    If err_msg = "" Then
        '成功
    ElseIf InStr(err_msg, "exit code=1") = 0 Then
        'exit code=1以外は上位に再度エラー通知
        Err.Raise 53, , "[DoPush] git pushでエラー (err_msg=" & err_msg & ")"
    Else
        'exit code=1は続行できる可能性が高いので確認
        If Common.ShowYesNoMessageBox( _
            "git pushで以下のエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[DoPush] git pushでエラー (err_msg=" & err_msg & ")"
        End If
    End If
    
    Common.WriteLog "DoPush E"
End Sub



