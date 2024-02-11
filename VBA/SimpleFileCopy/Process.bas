Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Public main_param As MainParam
Public sub_param As SubParam

Private target_files() As String
Private success_cnt As Long

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Erase target_files
    success_cnt = 0

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    'コピー先フォルダを削除する
    DeleteDestFolder
    
    Dim i As Long
    
    'メインループ
    For i = LBound(target_files) To UBound(target_files)
        'コピー
        CopyFiles target_files(i)
    Next i

    Common.WriteLog "Run E"
End Sub

Public Function GetResult() As String
    GetResult = success_cnt & "/" & sub_param.GetFilePathListCount()
End Function

'パラメータのチェックと収集を行う
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Set main_param = New MainParam
    Set sub_param = New SubParam
    main_param.Init
    sub_param.Init
    
    'Main Params
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()
    
    'Sub Params
    sub_param.Validate
    
    target_files = sub_param.GetFilePathList()

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'コピー先フォルダを削除する
Private Sub DeleteDestFolder()
    Common.WriteLog "DeleteDestFolder S"

    Dim dst_path As String: dst_path = main_param.GetDestDirPath()

    If Common.IsExistsFolder(dst_path) = True Then
        If Common.IsEmptyFolder(dst_path) = False Then
            If Common.ShowYesNoMessageBox( _
                "コピー先フォルダが空ではありません。" & vbCrLf & _
                "処理を続けますか？" & vbCrLf & _
                "（続けるとフォルダは削除されます!）" _
            ) = False Then
                Err.Raise 53, , "コピー先フォルダが空では無いので処理をキャンセルしました。(" & dst_path & ")"
            End If
        End If
    End If

    If Common.IsExistsFolder(dst_path) = True Then
        Common.DeleteFolder dst_path
    End If
    
    Common.CreateFolder dst_path

    Common.WriteLog "DeleteDestFolder E"
End Sub

'ファイルをコピーする
Private Sub CopyFiles(ByVal target_path As String)
    Common.WriteLog "CopyFiles S"
    Common.WriteLog "[CopyFiles] SrcFilePath=" & target_path
        
    Dim err_msg As String
    Dim dest_dir_path As String: dest_dir_path = main_param.GetDestDirPath()
    Dim is_copy_dir As Boolean: is_copy_dir = main_param.IsCopyDir()
    Dim is_continue As Boolean: is_continue = main_param.IsContinue()
    Dim is_overwrite As Boolean: is_overwrite = main_param.IsOverWrite
    Dim src_dir_path As String
    Dim dst_file_path As String
    
    If Common.IsExistsFile(target_path) = False Then
        err_msg = "コピー元ファイルが存在しません(" & target_path & ")"
        Common.WriteLog "[CopyFiles] ★Error! " & err_msg
        
        If is_continue = False Then
            'コピー元が存在しない場合は無視して続行しない
            Err.Raise 53, , err_msg
        End If
        
        Common.WriteLog "CopyFiles E-1"
        Exit Sub
    End If
    
    If is_copy_dir = False Then
        'フォルダはコピーしない場合
        
        'コピー先ファイルパスを作成
        dst_file_path = dest_dir_path & SEP & Common.GetFileName(target_path)
        
        Common.WriteLog "[CopyFiles] DestFilePath=" & target_path
        
        If Common.IsExistsFile(dst_file_path) = True Then
            'すでに同名ファイルがある場合
            
            If is_overwrite = False Then
                'コピー先に同名ファイルがある場合は上書きしない場合は、ユニークなファイル名に変更する
                Common.CopyUniqueFile target_path, dest_dir_path
                
                success_cnt = success_cnt + 1
                Common.WriteLog "CopyFiles E-2"
                Exit Sub
            End If
        End If
        
        '上書きコピー
        Common.CopyFile target_path, dst_file_path
        
        success_cnt = success_cnt + 1
        Common.WriteLog "CopyFiles E-3"
        Exit Sub
    End If
    
    'フォルダもコピーする場合
    
    'コピー先フォルダパスをコピー元から取得
    src_dir_path = Replace(Common.GetFolderPath(target_path), ":", "")
    
    'コピー先ファイルパスを作成
    dst_file_path = dest_dir_path & SEP & Replace(target_path, ":", "")
    
    Common.WriteLog "[CopyFiles] DestFilePath=" & dst_file_path
    
    'コピー
    Common.CopyFile target_path, dst_file_path, True
    
    success_cnt = success_cnt + 1
    Common.WriteLog "CopyFiles E"
End Sub
