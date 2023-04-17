VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Param_001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private external_path As String
Private dst_dir_path As String
Private is_debug_log As Boolean
Private is_debug_log_s As String
Private git_path As String
Private base_branch As String
Private target_list() As Param_002

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Param001] "

    external_path = ""
    dst_dir_path = ""
    is_debug_log = False
    git_path = ""
    base_branch = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    Const CLM = "D"
    Const i = 12
    
    external_path = main_sheet.Range(CLM & i).value
    dst_dir_path = main_sheet.Range(CLM & i + 1).value
    is_debug_log_s = main_sheet.Range(CLM & i + 2).value
    git_path = main_sheet.Range(CLM & i + 3).value
    base_branch = main_sheet.Range(CLM & i + 4).value

    CollectParam002
End Sub

Private Sub CollectParam002()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)

    Const CLM = "C"
    Const ROW = 20
    Dim i As Long: i = 0
    Dim path As String
    
    Do
        path = main_sheet.Range(CLM & ROW + i).value
        
        If path = "" Then
            Exit Do
        End If
        
        Dim param2 As Param_002
        Set param2 = New Param_002
        
        param2.Init ROW + i
        param2.Validate ROW + i
        
        ReDim Preserve target_list(i)
        Set target_list(i) = param2
        
        i = i + 1
    Loop
    
    If i = 0 Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが指定されていません(0個)"
    End If
    
End Sub

'検証
Public Sub Validate()
    If external_path = "" Then
        Err.Raise 53, , HEADER & "外部ツールファイルパスが指定されていません"
    End If
    
    If Common.IsExistsFile(external_path) = False Then
        Err.Raise 53, , HEADER & "外部ツールファイルパスが見つかりません (" & external_path & ")"
    End If
    
    If dst_dir_path = "" Then
        Err.Raise 53, , HEADER & "コピー先フォルダパスが指定されていません"
    End If
    
    If is_debug_log_s = "" Or is_debug_log_s = "NO" Then
        is_debug_log = False
    Else
        is_debug_log = True
    End If
    
    If git_path = "" Then
        Err.Raise 53, , HEADER & "Gitフォルダパスが指定されていません"
    End If
    
    If Common.IsExistsFolder(git_path) = False Then
        Err.Raise 53, , HEADER & "Gitフォルダパスが見つかりません (" & git_path & ")"
    End If

    If base_branch = "" Then
        Err.Raise 53, , HEADER & "ベースブランチ名が指定されていません"
    End If

End Sub

'=========================================
'Getter
'=========================================

Public Function GetExternalPath() As String
    GetExternalPath = external_path
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

Public Function IsDebugLog() As String
    IsDebugLog = is_debug_log
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

Public Function GetTargetList() As Param_002()
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

Public Function GetGitDirPath() As String
    GetGitDirPath = git_path
End Function

Public Function GetBaseBranch() As String
    GetBaseBranch = base_branch
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "external_path=(" & external_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_debug_log=(" & is_debug_log_s & "), " & _
          "git_path=(" & git_path & "), " & _
          "base_branch=(" & base_branch & ")"
          
    GetAllValue = ret
End Function

