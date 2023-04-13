Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String

'パラメータ
Private main_param As MainParam

'グローバル


'メイン処理
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator

    'パラメータのチェックと収集を行う
    If CheckAndCollectParam() = False Then
        Common.WriteLog "Run E1"
        Exit Sub
    End If

    '変換実行
    Convert
    
    Common.WriteLog "Run E"
    MsgBox "終わりました"
End Sub

'パラメータのチェックと収集を行う
Private Function CheckAndCollectParam() As Boolean
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    Set main_param = New MainParam
    err_msg = main_param.Init()
    If err_msg <> "" Then
        MsgBox err_msg
        CheckAndCollectParam = False
        Common.WriteLog "CheckAndCollectParam E1 (" & err_msg & ")"
        Exit Function
    End If
    
    Common.WriteLog main_param.GetAllValue()

    CheckAndCollectParam = True
    Common.WriteLog "CheckAndCollectParam E"
End Function

'変換を実施する
Private Sub Convert()
    Common.WriteLog "Convert S"
    
    If main_param.GetDestDirPath() = "" Then
        ConvertMain
    Else
        Common.DeleteFolder main_param.GetDestDirPath()
        Common.CopyFolder main_param.GetSrcDirPath(), main_param.GetDestDirPath()
        
        ConvertMain
    End If

    Common.WriteLog "Convert E"
End Sub

Private Sub ConvertMain()
    Dim i As Integer
    
    Dim src_file_list() As String
    Dim is_backup As Boolean
    
    If main_param.GetDestDirPath() = "" Then
        src_file_list = Common.CreateFileList(main_param.GetSrcDirPath(), main_param.GetExtension(), main_param.IsContainSubDir())
        is_backup = main_param.IsBackup()
    Else
        src_file_list = Common.CreateFileList(main_param.GetDestDirPath(), main_param.GetExtension(), main_param.IsContainSubDir())
        is_backup = False
    End If

    For i = LBound(src_file_list) To UBound(src_file_list)
        Dim path As String: path = src_file_list(i)

        If main_param.GetConvertType() = "SJIS→UTF8" Then
            Common.SJIStoUTF8 path, is_backup
        Else
            Common.UTF8toSJIS path, is_backup
        End If

        Common.WriteLog "i=" & i & ", path=" & path
    Next i
End Sub
