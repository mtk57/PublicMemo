VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamCommon"
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

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Common Param] "

    external_path = ""
    dst_dir_path = ""
    is_debug_log = False
End Sub

'初期化
Public Sub Init()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    Const CLM = Define.COMMON_PARAM_CLM
    Const i = Define.COMMON_PARAM_ROW
    
    external_path = sheet.Range(CLM & i).value
    dst_dir_path = sheet.Range(CLM & i + 1).value
    is_debug_log_s = sheet.Range(CLM & i + 2).value
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

Public Function IsDebugLog() As Boolean
    IsDebugLog = is_debug_log
End Function


Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "external_path=(" & external_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_debug_log=(" & is_debug_log_s & ")"
          
    GetAllValue = ret
End Function

