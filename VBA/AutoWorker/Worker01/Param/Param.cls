VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private common_param As ParamCommon
Private git_param As ParamGit
Private target_list() As ParamTarget

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Param] "
End Sub

'初期化
Public Sub Init()
    Set common_param = New ParamCommon
    Set git_param = New ParamGit
    
    common_param.Init
    git_param.Init
    CollectTarget
End Sub

Private Sub CollectTarget()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)

    Const CLM = Define.TARGET_PARAM_CLM_1
    Const row = Define.TARGET_PARAM_ROW
    Dim i As Long: i = 0
    Dim path As String
    
    Do
        path = main_sheet.Range(CLM & row + i).value
        
        If path = "" Then
            Exit Do
        End If
        
        Dim target As ParamTarget
        Set target = New ParamTarget
        
        target.Init row + i
        target.Validate row + i
        
        ReDim Preserve target_list(i)
        Set target_list(i) = target
        
        i = i + 1
    Loop
    
    If i = 0 Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが指定されていません(0個)"
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

Public Function GetCommonParam() As ParamCommon
    GetCommonParam = common_param
End Function

Public Function GetExternalPath() As String
    GetExternalPath = common_param.GetExternalPath()
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = common_param.GetDestDirPath()
End Function

Public Function IsDebugLog() As Boolean
    IsDebugLog = common_param.IsDebugLog()
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

Public Function GetBaseBranch() As String
    GetBaseBranch = git_param.GetBaseBranch()
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



