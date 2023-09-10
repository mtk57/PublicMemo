VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private proc_type As PROCESS_TYPE

Private common_param As ParamCommon
Private git_param As ParamGit
Private target_list() As ParamTarget

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[ParamContainer] "
    proc_type = UNKNOWN

    Set common_param = New ParamCommon
    Set git_param = New ParamGit
End Sub

'初期化
Public Sub Init()
    common_param.Init
    git_param.Init
    
    CollectTarget
End Sub

Private Sub CollectTarget()
    Dim params As Worksheet
    Set params = ThisWorkbook.Sheets("params")

    Const clm = Define.TARGET_PARAM_CLM_0
    Const row = Define.TARGET_PARAM_ROW
    Dim i As Long: i = 0
    Dim cnt As Long: cnt = 0
    Dim enable As String
    
    Do
        enable = params.Range(clm & row + i).value
        
        If enable = "" Then
            Exit Do
        End If
        
        Dim target As ParamTarget
        Set target = New ParamTarget
        
        target.Init row + i
        target.SetProcessType proc_type
        target.Validate row + i
        
        If target.IsEnable = True Then
            ReDim Preserve target_list(cnt)
            Set target_list(cnt) = target
            cnt = cnt + 1
        End If
        
        i = i + 1
    Loop
    
    If cnt = 0 Then
        Err.Raise 53, , HEADER & "有効なTarget行がありません。"
    End If
    
End Sub

'検証
Public Sub Validate()
    common_param.Validate
    git_param.Validate
End Sub

'=========================================
'Getter
'=========================================

Public Function GetProcessType() As PROCESS_TYPE
    GetProcessType = proc_type
End Function

Public Function GetCommonParam() As ParamCommon
    GetCommonParam = common_param
End Function

Public Function GetExternalPath() As String
    GetExternalPath = common_param.GetExternalPath()
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = common_param.GetDestDirPath()
End Function

Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = common_param.GetIgnoreFiles()
End Function

Public Function IsDebugLog() As Boolean
    IsDebugLog = common_param.IsDebugLog()
End Function

Public Function GetBatPath() As String
    GetBatPath = common_param.GetBatPath()
End Function

Public Function GetBatArgs() As String
    GetBatArgs = common_param.GetBatArgs()
End Function

Public Function GetGitParam() As ParamGit
    GetGitParam = git_param
End Function

Public Function GetUrl() As String
    GetUrl = git_param.GetUrl()
End Function

Public Function GetGitDirPath() As String
    GetGitDirPath = git_param.GetGitDirPath()
End Function

Public Function GetBaseFolder() As String
    GetBaseFolder = git_param.GetBaseFolder()
End Function

Public Function GetBaseBranch() As String
    GetBaseBranch = git_param.GetBaseBranch()
End Function

Public Function IsUpdateRemote() As Boolean
    IsUpdateRemote = git_param.IsUpdateRemote()
End Function

Public Function IsDeleteExistBranch() As Boolean
    IsDeleteExistBranch = git_param.IsDeleteExistBranch()
End Function

'-------------------------------------------------
'VBプロジェクト未参照ファイルを除外
'-------------------------------------------------
Public Function IsIgnoreNotRef() As Boolean
    IsIgnoreNotRef = git_param.IsIgnoreNotRef()
End Function

Public Function GetVBProjFilePathList() As String()
    Dim ret() As String
    Dim i As Long
    Dim cnt As Long: cnt = GetTargetListCount()
    ReDim Preserve ret(cnt - 1)
    For i = 0 To cnt - 1
        ret(i) = target_list(i).GetVBPrjFilePath()
    Next i
    GetVBProjFilePathList = ret
End Function

Public Function GetTargetList() As ParamTarget()
    GetTargetList = target_list
End Function

Public Function GetTargetListCount() As Long
    If Common.IsEmptyArray(target_list) = True Then
        GetTargetListCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(target_list)
    cnt = cnt + 1
    GetTargetListCount = cnt
End Function


Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          common_param.GetAllValue() & Chr(13) & _
          git_param.GetAllValue()
          
    GetAllValue = ret
End Function

'---------------------------
' Setter
'---------------------------
Public Sub SetProcessType(ByVal type_ As PROCESS_TYPE)
    proc_type = type_

    common_param.SetProcessType proc_type
    git_param.SetProcessType proc_type
End Sub


