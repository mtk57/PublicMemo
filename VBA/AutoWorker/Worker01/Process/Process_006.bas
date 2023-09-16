Attribute VB_Name = "Process_006"
Option Explicit

Private prms As ParamContainer
Private SEP As String
Private DQ As String

Private txtfile_num As Integer
Private is_txt_opened As Boolean
Private outpath As String

Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    Dim msg As String: msg = ""

    Set prms = New ParamContainer
    
    prms.SetProcessType PROCESS_TYPE.PROC_006
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    If Common.IsExistsFolder(prms.GetDestDirPath) = False Then
        Common.CreateFolder prms.GetDestDirPath()
    End If
    
    Dim i As Integer
    Dim target As ParamTarget
    Dim targetlist() As ParamTarget
    targetlist = prms.GetTargetList()
    
    outpath = prms.GetDestDirPath() & SEP & Common.GetNowTimeString() & ".txt"
    OpenTxt outpath
        
On Error GoTo ErrorHandler
    WorkerCommon.DoClone prms
    
    For i = LBound(targetlist) To UBound(targetlist)
    
        Set target = targetlist(i)
    
        WorkerCommon.SwitchDevelopBranch prms
        
        WorkerCommon.DoPull prms
        
        OutputTagList target
    
    Next i
        
    CloseTxt
    
    Common.WriteLog "Run E"
    
    Exit Sub

ErrorHandler:
    Dim err_msg As String: err_msg = Err.Description
    CloseTxt
    Common.WriteLog "Run E(Error)=" & err_msg
    Err.Raise 53, , "[Run] エラー! (err_msg=" & err_msg & ")"
End Sub

'タグ一覧を出力する
Private Sub OutputTagList(ByRef target As ParamTarget)
    Common.WriteLog "OutputTagList S"
    
    Dim tag_list() As String
    tag_list = GetTargetTagList(target)
    
    If Common.IsEmptyArray(tag_list) = True Then
        Common.WriteLog "Tag not found.(Branch=" & target.GetBranch() & ", Tag=" & target.GetTag() & ")"
        Common.WriteLog "OutputTagList E1"
        Exit Sub
    End If
    
    'タグ一覧をtxtで保存する
    Dim i As Long
    For i = LBound(tag_list) To UBound(tag_list)
        WriteTxt target.GetBranch() & vbTab & tag_list(i)
    Next i
        
    Common.WriteLog "OutputTagList E"
End Sub

Private Function GetTargetTagList(ByRef target As ParamTarget) As String()
    Common.WriteLog "GetTargetTagList S"

    Dim cmd As String
    Dim git_result() As String
    
    Dim find_word As String: find_word = target.GetTag()
    
    'タグを検索
    cmd = "git tag --list " & find_word
    git_result = Common.RunGit(prms.GetGitDirPath(), cmd)
    
    git_result = Common.DeleteEmptyArray(git_result)
    
    If Common.IsEmptyArray(git_result) = True Then
        Common.WriteLog "Tag not found.(" & find_word & ")"
    End If
    
    GetTargetTagList = git_result

    Common.WriteLog "GetTargetTagList E"
End Function

Private Sub OpenTxt(ByVal txtfile_path As String)
    If is_txt_opened = True Then
        'すでにオープンしているので無視
        Exit Sub
    End If
    txtfile_num = FreeFile()
    Open txtfile_path For Append As txtfile_num
    is_txt_opened = True
End Sub

Private Sub WriteTxt(ByVal contents As String)
    If is_txt_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Print #txtfile_num, contents
    'Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & contents
End Sub

Private Sub CloseTxt()
    If is_txt_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Close txtfile_num
    txtfile_num = -1
    is_txt_opened = False
End Sub

